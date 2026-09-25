import { formatWith, type CellFormat } from '../sheet/Format';
import { Formats } from '../sheet/Formats';
import { Sheet } from '../sheet/Sheet';

/**
 * One change to one cell, and what it replaced.
 *
 * Two kinds rather than one since Phase 9, because formatting a range
 * and typing into it are both things a person undoes and neither is
 * the other. Keeping them in the same list is what makes a paste that
 * carried formats one press of ctrl-Z rather than two.
 */
type Edit = TextEdit | FormatEdit | RegionEdit;

interface TextEdit {
  readonly kind: 'text';
  readonly row: number;
  readonly column: number;
  readonly before: string;
  readonly after: string;
}

/** A cell's palette id, before and after. Ids, not formats: the
 * palette only grows, so an id undoes to an entry that is still
 * there, and a step stays four numbers however large the format is. */
interface FormatEdit {
  readonly kind: 'format';
  readonly row: number;
  readonly column: number;
  readonly before: number;
  readonly after: number;
}

/**
 * A whole row, column or sheet formatted at once.
 *
 * One entry however many cells it covers, which is the point: a
 * million cell edits for one press of ctrl-B is a step nobody can
 * afford to keep and a file nobody asked for.
 *
 * `overrides` is the cells inside the region that had a format of
 * their own, before and after — the change is applied to those too,
 * so that making a sheet bold keeps the currency in column C. It is
 * bounded by how many cells anybody has actually formatted.
 */
interface RegionEdit {
  readonly kind: 'region';
  readonly scope: 'sheet' | 'row' | 'column';
  readonly index: number;
  readonly row: number;
  readonly column: number;
  readonly before: number;
  readonly after: number;
  readonly overrides: readonly { readonly key: number; readonly before: number; readonly after: number }[];
}

/**
 * What one press of ctrl-Z takes back.
 *
 * A step rather than an edit, because a paste is one action and fifty
 * cells: undoing it a cell at a time would be fifty presses to
 * reverse one, which is the kind of thing that makes people stop
 * trusting undo.
 */
type Step = readonly Edit[];

/**
 * The sheet plus the two things a document has that a model does not:
 * where the selection is, and what can be taken back.
 *
 * Still plain TypeScript — no RxJS, no framework — so it runs in node
 * beside `Sheet` and is the layer a spec drives when it wants to ask
 * about behaviour rather than about the wire.
 *
 * Undo is per commit, not per keystroke. Phase 4 brings coalescing
 * when there is a cell editor to coalesce, and `EditableTextModel`
 * already has its own; what belongs here is the commit, which is the
 * unit a person means when they press ctrl-Z in a spreadsheet.
 */
export class SheetDocument {
  readonly sheet = new Sheet();
  readonly formats = new Formats();

  private readonly undoStack: Step[] = [];
  private readonly redoStack: Step[] = [];
  /** Edits collected by an open `transact`, or null outside one. */
  private collecting: Edit[] | null = null;

  selection = { row: 0, column: 0, anchorRow: 0, anchorColumn: 0 };

  get canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  /**
   * Commits a cell, recording what it replaced.
   *
   * A commit that changes nothing is not an edit: retyping the same
   * text, or pressing Enter on a cell without touching it, must not
   * put an entry on the stack that undoes to itself and looks broken.
   */
  setCell(row: number, column: number, input: string): void {
    const before = this.sheet.input(row, column);
    if (before === input) {
      return;
    }
    this.writeCell(row, column, input);
    this.record({ kind: 'text', row, column, before, after: input });
  }

  /**
   * The write itself, with the cell's own format consulted.
   *
   * Separate from `setCell` because undo and redo write without
   * recording, and all three have to agree about what a Text-
   * formatted cell does with `007`.
   */
  private writeCell(row: number, column: number, input: string): void {
    this.sheet.setCell(row, column, input, this.formats.formatAt(row, column).number.kind === 'text');
  }

  /**
   * Formats a cell, recording the palette id it replaced.
   *
   * A format that changes nothing is not an edit, on the same rule as
   * a commit that changes nothing: formatting a bold cell bold again
   * must not put an entry on the stack that undoes to itself.
   *
   * Re-writing the cell afterwards is not busywork. A format can
   * change what its *value* is — Text is the one that does — so a
   * cell holding the number seven and formatted as Text has to become
   * a cell holding `7` as text, and the only way to get there is
   * through the same path that typing takes.
   */
  setFormat(row: number, column: number, format: CellFormat): void {
    const before = this.formats.idAt(row, column);
    const after = this.formats.idFor(format);
    if (before === after) {
      return;
    }
    this.applyFormat(row, column, after);
    this.record({ kind: 'format', row, column, before, after });
  }

  /**
   * Formats a whole row, column or sheet.
   *
   * The command layer decides when a selection is a region — see
   * `SheetService.format` — and this is what it calls. The saving is
   * not in the writing but in what is *kept*: one undo entry, one
   * number in the file, and one lookup per cell rather than an entry
   * per cell in all three.
   */
  formatRegion(scope: 'sheet' | 'row' | 'column', index: number, change: (format: CellFormat) => CellFormat): void {
    const heldId = scope === 'sheet' ? this.formats.sheetId : scope === 'row' ? this.formats.rowId(index) : this.formats.columnId(index);
    const after = this.formats.idFor(change(this.formats.byId(heldId)));

    // The change applies to the cells inside the region that had a
    // format of their own as well, so that making the sheet bold
    // keeps the currency somebody put in column C.
    const overrides = this.formats.overridesIn(scope, index).map(([key, before]) => ({
      key,
      before,
      after: this.formats.idFor(change(this.formats.byId(before)))
    }));

    if (after === heldId && overrides.every(entry => entry.after === entry.before)) {
      return;
    }
    this.writeRegion(scope, index, after, overrides.map(entry => ({ key: entry.key, id: entry.after })));
    this.record({
      kind: 'region',
      scope,
      index,
      row: scope === 'row' ? index : 0,
      column: scope === 'column' ? index : 0,
      before: heldId,
      after,
      overrides
    });
  }

  /** Writes a region and its overrides, and re-reads what it covers. */
  private writeRegion(
    scope: 'sheet' | 'row' | 'column',
    index: number,
    id: number,
    overrides: readonly { key: number; id: number }[]
  ): void {
    this.formats.setRegion(scope, index, id);
    for (const entry of overrides) {
      this.formats.setIdByKey(entry.key, entry.id);
    }
    this.retextRegion(scope, index);
  }

  /**
   * Re-reads every cell a region covers, in case its Text-ness moved.
   *
   * Bounded by the cells that hold something, not by the region: a
   * sheet-wide format walks the store, which is the number of cells
   * somebody has actually typed in.
   */
  private retextRegion(scope: 'sheet' | 'row' | 'column', index: number): void {
    for (const cell of this.sheet.entries()) {
      if (scope === 'row' && cell.row !== index) {
        continue;
      }
      if (scope === 'column' && cell.column !== index) {
        continue;
      }
      if (cell.input !== '') {
        this.writeCell(cell.row, cell.column, cell.input);
      }
    }
  }

  private applyFormat(row: number, column: number, id: number): void {
    const wasText = this.formats.formatAt(row, column).number.kind === 'text';
    this.formats.setId(row, column, id);
    const isText = this.formats.byId(id).number.kind === 'text';
    if (wasText !== isText) {
      const input = this.sheet.input(row, column);
      if (input !== '') {
        this.writeCell(row, column, input);
      }
    }
  }

  /** A cell's format, as the palette holds it. */
  formatAt(row: number, column: number): CellFormat {
    return this.formats.formatAt(row, column);
  }

  /**
   * What the screen shows for a cell: its value under its format.
   *
   * Moved off `Sheet` in Phase 9, because a display string now
   * depends on two things the sheet does not own both of. It stays on
   * *this* thread either way — the render worker receives the string
   * and never the format that made it, which is what keeps a
   * locale-aware number formatter off the frame path.
   */
  display(row: number, column: number): string {
    return formatWith(this.sheet.value(row, column), this.formats.formatAt(row, column).number);
  }

  /**
   * Runs several edits as one step.
   *
   * A paste, a cut and a fill are each one action to the person doing
   * them, so they are one entry on the stack however many cells they
   * touched. Nested calls join the step already open rather than
   * opening another, so a fill that pastes is still one press of
   * ctrl-Z.
   */
  transact(run: () => void): void {
    if (this.collecting !== null) {
      run();
      return;
    }
    const step: Edit[] = [];
    this.collecting = step;
    try {
      run();
    } finally {
      this.collecting = null;
    }
    if (step.length > 0) {
      this.undoStack.push(step);
      this.redoStack.length = 0;
    }
  }

  private record(edit: Edit): void {
    if (this.collecting !== null) {
      this.collecting.push(edit);
      return;
    }
    this.undoStack.push([edit]);
    // A new edit is a new future; whatever was undone is unreachable.
    this.redoStack.length = 0;
  }

  undo(): boolean {
    const step = this.undoStack.pop();
    if (step === undefined) {
      return false;
    }
    // Backwards: two edits to one cell in one step have to be undone
    // in the order they were made or the earlier one wins.
    for (let at = step.length - 1; at >= 0; at--) {
      const edit = step[at];
      if (edit.kind === 'text') {
        this.writeCell(edit.row, edit.column, edit.before);
      } else if (edit.kind === 'region') {
        this.writeRegion(
          edit.scope,
          edit.index,
          edit.before,
          edit.overrides.map(entry => ({ key: entry.key, id: entry.before }))
        );
      } else {
        this.applyFormat(edit.row, edit.column, edit.before);
      }
    }
    this.redoStack.push(step);
    this.selectStep(step);
    return true;
  }

  redo(): boolean {
    const step = this.redoStack.pop();
    if (step === undefined) {
      return false;
    }
    for (const edit of step) {
      if (edit.kind === 'text') {
        this.writeCell(edit.row, edit.column, edit.after);
      } else if (edit.kind === 'region') {
        this.writeRegion(
          edit.scope,
          edit.index,
          edit.after,
          edit.overrides.map(entry => ({ key: entry.key, id: entry.after }))
        );
      } else {
        this.applyFormat(edit.row, edit.column, edit.after);
      }
    }
    this.undoStack.push(step);
    this.selectStep(step);
    return true;
  }

  /** Takes the selection to what a step changed, so it is seen. */
  private selectStep(step: Step): void {
    let firstRow = Number.POSITIVE_INFINITY;
    let lastRow = Number.NEGATIVE_INFINITY;
    let firstColumn = Number.POSITIVE_INFINITY;
    let lastColumn = Number.NEGATIVE_INFINITY;
    for (const edit of step) {
      firstRow = Math.min(firstRow, edit.row);
      lastRow = Math.max(lastRow, edit.row);
      firstColumn = Math.min(firstColumn, edit.column);
      lastColumn = Math.max(lastColumn, edit.column);
    }
    this.selection = { row: firstRow, column: firstColumn, anchorRow: lastRow, anchorColumn: lastColumn };
  }

  setSelection(row: number, column: number, anchorRow: number, anchorColumn: number): void {
    this.selection = { row, column, anchorRow, anchorColumn };
  }

  /** Every non-empty cell, for a repository or a search. */
  *entries(): Generator<{ row: number; column: number; input: string }> {
    yield* this.sheet.entries();
  }

  /** What the formula bar shows for the active cell. */
  get activeInput(): string {
    return this.sheet.input(this.selection.row, this.selection.column);
  }
}
