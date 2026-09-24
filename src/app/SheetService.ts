import { BehaviorSubject, type Observable } from 'rxjs';

import { ROW_HEIGHT, COLUMN_WIDTH } from './dimensions';
import {
  EMPTY_WINDOW,
  type SheetClipboard,
  type SheetEditor,
  type SheetGeometry,
  type SheetSelection,
  type SheetStatus,
  type SheetWindow
} from './SheetContract';
import type { SheetDocument } from './SheetDocument';
import { clearRect, copyRect, fillRect, fillTarget, pasteBlock, rectOf, type CopyOrigin } from './SheetRanges';

/**
 * Runs a continuation later, as a task rather than a microtask.
 *
 * It has to be a *task*. A microtask runs before the thread returns to
 * its message queue, so a recalc sliced across microtasks never lets a
 * `setViewport` command in and is exactly the uninterrupted recalc the
 * budget exists to avoid — the sheet would still go blank, and the
 * slicing would look like it was working.
 */
export type Schedule = (run: () => void) => void;

const defaultSchedule: Schedule = run => {
  setTimeout(run, 0);
};

export interface SheetServiceOptions {
  /** Cells evaluated per slice before the thread is handed back. */
  readonly budget?: number;
  readonly schedule?: Schedule;
  readonly rowCount?: number;
  readonly columnCount?: number;
}

/**
 * The application worker's view of the sheet, shaped for the wire.
 *
 * This is the last layer before the barrier and the only one that has
 * to care that plain data is all that crosses. Everything below it —
 * the document, the sheet, the parser, the graph — is unaware there is
 * a barrier at all.
 *
 * **The recalc pump is the point of this phase.** An edit marks work
 * and the pump does it a slice at a time, publishing the window after
 * each slice and handing the thread back in between. Phase 0 measured
 * what happens without it: thirty milliseconds of uninterrupted
 * application thread leaves the sheet blank in 89% of frames, because
 * the render worker goes on scrolling at 60fps and asking for windows
 * that nobody is free to serve. A `setViewport` arriving mid-recalc is
 * answered on the spot, ahead of the arithmetic.
 */
export class SheetService {
  readonly window: Observable<SheetWindow>;
  readonly geometry: Observable<SheetGeometry>;
  readonly selection: Observable<SheetSelection>;
  readonly editor: Observable<SheetEditor>;
  readonly status: Observable<SheetStatus>;
  readonly clipboard: Observable<SheetClipboard>;

  /** Slices run, for a spec that wants to know the pump ran at all. */
  readonly stats = { slices: 0, publishes: 0 };

  private readonly windowSubject = new BehaviorSubject<SheetWindow>(EMPTY_WINDOW);
  private readonly geometrySubject: BehaviorSubject<SheetGeometry>;
  private readonly selectionSubject: BehaviorSubject<SheetSelection>;
  private readonly editorSubject: BehaviorSubject<SheetEditor>;
  private readonly statusSubject: BehaviorSubject<SheetStatus>;
  private readonly clipboardSubject = new BehaviorSubject<SheetClipboard>({ text: '', serial: 0 });
  /**
   * What this sheet last copied, and from where.
   *
   * Kept so that pasting it back knows how far it moved. Text from
   * anywhere else has no origin and is written as it arrived: a block
   * out of Excel means what it says, and moving references that were
   * never relative to this sheet would be inventing an intent.
   */
  private copied: CopyOrigin | null = null;
  private serial = 0;

  private viewport = { firstRow: 0, lastRow: -1, firstColumn: 0, lastColumn: -1 };
  private readonly budget: number;
  private readonly schedule: Schedule;
  private pumping = false;

  constructor(
    private readonly document: SheetDocument,
    options: SheetServiceOptions = {}
  ) {
    this.budget = options.budget ?? 2_000;
    this.schedule = options.schedule ?? defaultSchedule;
    this.geometrySubject = new BehaviorSubject<SheetGeometry>({
      rowCount: options.rowCount ?? 10_000,
      columnCount: options.columnCount ?? 100,
      rowHeight: ROW_HEIGHT,
      columnWidth: COLUMN_WIDTH
    });
    this.selectionSubject = new BehaviorSubject<SheetSelection>(document.selection);
    this.editorSubject = new BehaviorSubject<SheetEditor>({
      row: document.selection.row,
      column: document.selection.column,
      input: document.activeInput
    });
    this.statusSubject = new BehaviorSubject<SheetStatus>(this.statusNow());

    this.window = this.windowSubject;
    this.geometry = this.geometrySubject;
    this.selection = this.selectionSubject;
    this.editor = this.editorSubject;
    this.status = this.statusSubject;
    this.clipboard = this.clipboardSubject;
  }

  // ---------------------------------------------------------------------
  // Commands
  // ---------------------------------------------------------------------

  /**
   * The render worker's window moved.
   *
   * Published at once, whatever else is happening. This is the line
   * that keeps a scroll answerable during a recalc: the cells for the
   * new range are read straight out of the store, which holds correct
   * values for everything the recalc has already reached and stale
   * ones for what it has not — and a stale value on screen for two
   * frames is not a blank sheet for two seconds.
   */
  setViewport(firstRow: number, lastRow: number, firstColumn: number, lastColumn: number): void {
    this.viewport = { firstRow, lastRow, firstColumn, lastColumn };
    this.publishWindow();
  }

  setCell(row: number, column: number, input: string): void {
    this.document.setCell(row, column, input);
    // The edited cell's own value is settled already — a literal is
    // itself and a formula is queued — so the window can go out before
    // any arithmetic, which is what makes typing feel immediate.
    this.publishWindow();
    this.publishEditor();
    this.publishStatus();
    this.pump();
  }

  setSelection(row: number, column: number, anchorRow: number, anchorColumn: number): void {
    this.document.setSelection(row, column, anchorRow, anchorColumn);
    this.selectionSubject.next(this.document.selection);
    this.publishEditor();
  }

  undo(): void {
    if (this.document.undo()) {
      this.afterHistory();
    }
  }

  redo(): void {
    if (this.document.redo()) {
      this.afterHistory();
    }
  }

  copy(cut: boolean): void {
    const rect = rectOf(this.document.selection);
    const text = copyRect(this.document, rect);
    this.copied = { text, row: rect.firstRow, column: rect.firstColumn };
    this.serial++;
    this.clipboardSubject.next({ text, serial: this.serial });
    if (cut) {
      clearRect(this.document, rect);
      this.afterEdit();
    }
  }

  paste(text: string): void {
    const at = rectOf(this.document.selection);
    const written = pasteBlock(this.document, text, { row: at.firstRow, column: at.firstColumn }, this.copied);
    this.document.setSelection(written.firstRow, written.firstColumn, written.lastRow, written.lastColumn);
    this.selectionSubject.next(this.document.selection);
    this.afterEdit();
  }

  clearRange(): void {
    clearRect(this.document, rectOf(this.document.selection));
    this.afterEdit();
  }

  fill(toRow: number, toColumn: number): void {
    const source = rectOf(this.document.selection);
    const target = fillTarget(source, toRow, toColumn);
    fillRect(this.document, source, target);
    this.document.setSelection(target.firstRow, target.firstColumn, target.lastRow, target.lastColumn);
    this.selectionSubject.next(this.document.selection);
    this.afterEdit();
  }

  /** What every edit that is not a single keystroke has to do afterwards. */
  private afterEdit(): void {
    this.publishWindow();
    this.publishEditor();
    this.publishStatus();
    this.pump();
  }

  private afterHistory(): void {
    this.selectionSubject.next(this.document.selection);
    this.publishWindow();
    this.publishEditor();
    this.publishStatus();
    this.pump();
  }

  // ---------------------------------------------------------------------
  // The pump
  // ---------------------------------------------------------------------

  /** True while a recalc is still in slices. */
  get recalculating(): boolean {
    return this.pumping;
  }

  private pump(): void {
    if (this.pumping || this.document.sheet.pending === 0) {
      return;
    }
    this.pumping = true;
    this.step();
  }

  private step(): void {
    const result = this.document.sheet.recalculate(this.budget);
    this.stats.slices++;
    this.publishWindow();
    this.publishStatus();
    if (result.done) {
      this.pumping = false;
      return;
    }
    this.schedule(() => this.step());
  }

  // ---------------------------------------------------------------------
  // Publishing
  // ---------------------------------------------------------------------

  /**
   * The window, rebuilt from the store.
   *
   * Rebuilt whole rather than patched, because the differ is what
   * decides the wire: two structurally equal windows produce no
   * patches at all and `provide` sends nothing, so publishing after
   * every slice costs a walk of the visible cells and nothing else.
   * The walk is a few hundred string comparisons; the alternative is
   * bookkeeping that has to be right about which cells a recalc
   * touched, which is the same information the differ already has.
   */
  private publishWindow(): void {
    const { firstRow, lastRow, firstColumn, lastColumn } = this.viewport;
    if (lastRow < firstRow || lastColumn < firstColumn) {
      this.windowSubject.next(EMPTY_WINDOW);
      return;
    }
    const cells: Record<string, Record<string, string>> = {};
    for (let row = firstRow; row <= lastRow; row++) {
      const line: Record<string, string> = {};
      for (let column = firstColumn; column <= lastColumn; column++) {
        line[column] = this.document.sheet.display(row, column);
      }
      cells[row] = line;
    }
    this.stats.publishes++;
    this.windowSubject.next({ firstRow, lastRow, firstColumn, lastColumn, cells });
  }

  private publishEditor(): void {
    const { row, column } = this.document.selection;
    this.editorSubject.next({ row, column, input: this.document.activeInput });
  }

  private publishStatus(): void {
    this.statusSubject.next(this.statusNow());
  }

  private statusNow(): SheetStatus {
    return {
      pending: this.document.sheet.pending,
      canUndo: this.document.canUndo,
      canRedo: this.document.canRedo
    };
  }
}
