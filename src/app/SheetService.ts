import { BehaviorSubject, type Observable } from 'rxjs';

import { ROW_HEIGHT, COLUMN_WIDTH, MIN_COLUMN_WIDTH } from './dimensions';
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
import { columnName } from '../sheet/A1';
import { snapshotOf, applySnapshot, type SheetSnapshot } from './SheetFile';
import type { SheetRepository } from './SheetRepository';
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
  /** Where the sheet is kept. Without one it is kept nowhere. */
  readonly repository?: SheetRepository;
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
  private readonly repository: SheetRepository | undefined;
  private pumping = false;
  /**
   * Whether a load has finished.
   *
   * Saving before it has would write an empty sheet over a real one:
   * the seed runs first, the file arrives second, and between them the
   * document is neither.
   */
  private restored = false;
  /** How long the stress chain is, so it is built once. */
  private stressCells = 0;
  private stressRuns = 0;

  constructor(
    private readonly document: SheetDocument,
    options: SheetServiceOptions = {}
  ) {
    this.budget = options.budget ?? 2_000;
    this.schedule = options.schedule ?? defaultSchedule;
    const columnCount = options.columnCount ?? 100;
    this.repository = options.repository;
    this.geometrySubject = new BehaviorSubject<SheetGeometry>({
      rowCount: options.rowCount ?? 10_000,
      columnCount,
      rowHeight: ROW_HEIGHT,
      columnWidth: COLUMN_WIDTH,
      columnWidths: Array.from({ length: columnCount }, () => COLUMN_WIDTH)
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
    this.persist();
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

  setColumnWidth(column: number, width: number): void {
    const geometry = this.geometrySubject.value;
    if (column < 0 || column >= geometry.columnCount) {
      return;
    }
    const columnWidths = [...geometry.columnWidths];
    columnWidths[column] = Math.max(MIN_COLUMN_WIDTH, Math.round(width));
    this.geometrySubject.next({ ...geometry, columnWidths });
    this.persist();
  }

  /**
   * A chain of `cells` cells, each reading the one before it.
   *
   * Laid out across the sheet in rows so that a chain longer than the
   * sheet is tall still fits, and written through the model rather
   * than the document so that two hundred thousand cells are not two
   * hundred thousand undo entries. Built once; every call after moves
   * the head, which is what makes the whole chain out of date.
   *
   * It starts one row past the end of the sheet, so every cell in it
   * is somewhere nobody can scroll to, select, or type in. Two reasons,
   * and the second was found by reloading the page rather than by
   * reading the code: the claim is that recalculating cells nobody can
   * see costs the scroll nothing, so the cells have to be ones nobody
   * can see — and a cell outside the sheet is not part of the document,
   * which is what keeps a quarter of a million formulas from being
   * saved to somebody's file the first time they press the button. See
   * `snapshotOf`.
   */
  stress(cells: number): void {
    const sheet = this.document.sheet;
    const { columnCount, rowCount } = this.geometrySubject.value;
    if (this.stressCells !== cells) {
      const first = rowCount;
      sheet.setCell(first, 0, '1');
      // `cells` formulas, plus the literal head they hang from, so the
      // number on the button is the number of cells that go out of
      // date rather than the number of cells written.
      for (let at = 1; at <= cells; at++) {
        const row = first + Math.floor(at / columnCount);
        const column = at % columnCount;
        const fromRow = first + Math.floor((at - 1) / columnCount);
        const fromColumn = (at - 1) % columnCount;
        sheet.setCell(row, column, `=${columnName(fromColumn)}${fromRow + 1}+1`);
      }
      this.stressCells = cells;
    }
    this.stressRuns++;
    sheet.setCell(rowCount, 0, String(this.stressRuns));
    this.publishWindow();
    this.publishStatus();
    this.pump();
  }

  // ---------------------------------------------------------------------
  // Keeping it
  // ---------------------------------------------------------------------

  /**
   * Loads what was stored, or seeds a sheet that has never been opened.
   *
   * Async, and called after `serveChannels` rather than before it: a
   * channel served late misses the handshake, and the render worker is
   * perfectly able to draw an empty grid for the frame it takes to
   * read a file. The seed is written straight back, so what is on disk
   * from the second run onwards is a file this build wrote.
   */
  async restore(seed?: (document: SheetDocument) => void): Promise<void> {
    const stored = (await this.repository?.load()) ?? null;
    if (stored === null) {
      seed?.(this.document);
      this.document.sheet.recalculate();
    } else {
      applySnapshot(this.document, stored);
      this.geometrySubject.next({ ...this.geometrySubject.value, columnWidths: [...stored.columnWidths] });
    }
    this.restored = true;
    this.selectionSubject.next(this.document.selection);
    this.publishWindow();
    this.publishEditor();
    this.publishStatus();
    if (stored === null) {
      this.persist();
    }
  }

  /** The snapshot as it stands, for a spec or a worker shutting down. */
  snapshot(): SheetSnapshot {
    const { columnWidths, rowCount } = this.geometrySubject.value;
    return snapshotOf(this.document, columnWidths, rowCount);
  }

  /** Writes anything outstanding now. */
  flush(): Promise<void> {
    return this.repository?.flush() ?? Promise.resolve();
  }

  private persist(): void {
    if (!this.restored) {
      return;
    }
    this.repository?.save(this.snapshot());
  }

  /** What every edit that is not a single keystroke has to do afterwards. */
  private afterEdit(): void {
    this.publishWindow();
    this.publishEditor();
    this.publishStatus();
    this.persist();
    this.pump();
  }

  private afterHistory(): void {
    this.selectionSubject.next(this.document.selection);
    this.publishWindow();
    this.publishEditor();
    this.publishStatus();
    // An undo is an edit as far as the file is concerned. Left out,
    // taking something back and closing the tab would bring it back on
    // the next open, which is the opposite of what undo promises.
    this.persist();
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
      evaluated: this.document.sheet.stats.evaluated,
      canUndo: this.document.canUndo,
      canRedo: this.document.canRedo
    };
  }
}
