import { parseTypedDate } from '../sheet/Dates';
import { formatWith, type CellFormat, type NumberFormat } from '../sheet/Format';
import { Formats } from '../sheet/Formats';
import { Merges } from '../sheet/Merges';
import { Sheet } from '../sheet/Sheet';
import { shiftIndex, type Shift } from '../sheet/Shift';

/**
 * One change to one cell, and what it replaced.
 *
 * Two kinds rather than one since Phase 9, because formatting a range
 * and typing into it are both things a person undoes and neither is
 * the other. Keeping them in the same list is what makes a paste that
 * carried formats one press of ctrl-Z rather than two.
 */
type Edit = TextEdit | FormatEdit | RegionEdit | StructureEdit;

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
 * A row or column inserted or deleted.
 *
 * Recorded as the shift plus what the shift destroyed, because a
 * shift is not reversible on its own: deleting a row takes the cells
 * in it away and turns every reference to them into `#REF!`, and
 * neither comes back from applying the opposite shift. So the cells
 * that were removed and the formulas that were rewritten are kept, at
 * the positions they had *before* — which is where the inverse shift
 * puts everything back.
 *
 * `rewritten` is proportional to the formulas that actually mentioned
 * the line, not to the sheet. On a normal sheet that is tens; on a
 * column of fifty thousand chained formulas it is fifty thousand,
 * which is how much information the edit really changed and the price
 * of being able to take it back.
 */
interface StructureEdit {
  readonly kind: 'structure';
  readonly row: number;
  readonly column: number;
  readonly shift: Shift;
  /** Cells the shift removed, at the positions they had. */
  readonly removed: readonly { readonly row: number; readonly column: number; readonly input: string }[];
  /** Formulas the shift changed, as they read before it. */
  readonly rewritten: readonly { readonly row: number; readonly column: number; readonly input: string }[];
  /** Column widths as they were, so undo puts them back. */
  readonly widths: readonly number[];
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
  /**
   * How wide each column is drawn.
   *
   * Here rather than on the service's geometry since Phase 10, and
   * the reason is undo. A column insert moves the widths along with
   * the columns they describe, and taking that back has to move them
   * back — so the widths have to be somewhere the undo stack can
   * reach. Phase 6 had already made them the document's as far as the
   * *file* was concerned; this finishes the move.
   */
  columnWidths: number[] = [];
  /**
   * The rows somebody has hidden.
   *
   * A set of exceptions rather than a height per row, for the reason
   * `UiVirtualSheetOptions.rowHeights` is sparse: a sheet is ten
   * thousand rows tall and all but a handful are the same.
   */
  readonly hiddenRows = new Set<number>();
  /**
   * Rows a filter is hiding, kept apart from the ones somebody hid.
   *
   * Two sets rather than one because they are undone by different
   * things: clearing a filter must not reveal a row that was hidden
   * on purpose, and showing a hidden row must not fight the filter
   * that is hiding it. What the screen sees is the union.
   */
  readonly filteredRows = new Set<number>();
  /** How many rows and columns stay put while the rest scrolls. */
  frozenRows = 0;
  frozenColumns = 0;
  readonly merges = new Merges();

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
    // One step, because typing a date is one action: it writes a
    // serial number and the format that makes the serial legible, and
    // undoing it has to take both back.
    this.transact(() => {
      this.writeCell(row, column, input);
      this.record({ kind: 'text', row, column, before, after: input });
      this.formatTypedDate(row, column, input);
    });
  }

  /**
   * A typed date brings its format with it.
   *
   * The engine turns `2026-09-24` into 46,289 — see `literalValue` —
   * and on its own that is a cell showing forty-six thousand to
   * somebody who typed a date. The format is what finishes the job,
   * and it is applied here because the format axis is the document's
   * and the value is the sheet's.
   *
   * Only over `General`, and that is the whole of the rule. A cell
   * somebody has deliberately formatted has been answered already: a
   * date typed into a currency column shows as currency, which looks
   * odd and is what every spreadsheet does, because the alternative is
   * a format that silently undoes a choice somebody made on purpose.
   * The pattern is the one they typed in, so slashes give slashes
   * back.
   */
  private formatTypedDate(row: number, column: number, input: string): void {
    const current = this.formats.formatAt(row, column);
    if (current.number.kind !== 'general') {
      return;
    }
    const typed = parseTypedDate(input);
    if (typed === null) {
      return;
    }
    const number: NumberFormat =
      typed.date === null
        ? { kind: 'time', pattern: typed.time ?? 'hm' }
        : typed.time === null
          ? { kind: 'date', pattern: typed.date }
          : { kind: 'datetime', date: typed.date, time: typed.time };
    this.setFormat(row, column, { ...current, number });
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
      } else if (edit.kind === 'structure') {
        this.undoShift(edit);
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
      } else if (edit.kind === 'structure') {
        this.sheet.shift(edit.shift);
        this.formats.shift(edit.shift);
        this.columnWidths = shiftWidths(edit.widths, edit.shift);
      } else {
        this.applyFormat(edit.row, edit.column, edit.after);
      }
    }
    this.undoStack.push(step);
    this.selectStep(step);
    return true;
  }

  /**
   * Puts a structural change back.
   *
   * The opposite shift first, which restores every position, and then
   * the two things a shift destroys: the cells that were in a deleted
   * row, and the formulas it turned into `#REF!`. Both are written at
   * the positions they had before, which is where the opposite shift
   * has just put everything else.
   */
  private undoShift(edit: StructureEdit): void {
    this.sheet.shift({ ...edit.shift, by: -edit.shift.by });
    this.formats.shift({ ...edit.shift, by: -edit.shift.by });
    this.columnWidths = [...edit.widths];
    for (const cell of edit.removed) {
      this.writeCell(cell.row, cell.column, cell.input);
    }
    for (const cell of edit.rewritten) {
      this.writeCell(cell.row, cell.column, cell.input);
    }
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

  /**
   * Inserts or deletes rows or columns, as one step.
   *
   * The whole of it is one entry on the undo stack, because it is one
   * action to the person doing it however far it reached — the same
   * rule `transact` exists for.
   *
   * The column widths move with the columns they describe: inserting
   * a column in front of a wide one and leaving the widths alone
   * makes the wrong column wide.
   */
  applyShift(shift: Shift): void {
    const removed: { row: number; column: number; input: string }[] = [];
    const rewritten: { row: number; column: number; input: string }[] = [];
    for (const cell of this.sheet.entries()) {
      const index = shift.axis === 'row' ? cell.row : cell.column;
      if (shiftIndex(index, shift) === -1) {
        removed.push({ ...cell });
      } else if (cell.input.startsWith('=')) {
        // Kept whether or not it changes. Deciding here would mean
        // shifting every formula twice, and the list is dropped by
        // `record` if the step turns out to be empty anyway.
        rewritten.push({ ...cell });
      }
    }

    const widths = [...this.columnWidths];
    const hidden = [...this.hiddenRows];
    const filtered = [...this.filteredRows];
    this.sheet.shift(shift);
    this.formats.shift(shift);
    this.merges.shift(shift);
    this.columnWidths = shiftWidths(this.columnWidths, shift);
    // A hidden row is hidden by index, so it moves with the rows it
    // was among — an insert above a hidden row must not reveal it and
    // hide its neighbour instead.
    if (shift.axis === 'row') {
      shiftRows(this.hiddenRows, hidden, shift);
      shiftRows(this.filteredRows, filtered, shift);
    }

    this.record({
      kind: 'structure',
      row: shift.axis === 'row' ? shift.at : 0,
      column: shift.axis === 'column' ? shift.at : 0,
      shift,
      removed,
      rewritten,
      widths
    });
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

/**
 * Column widths, moved by a column insert or delete.
 *
 * The array stays the same length — the sheet is as wide as it was —
 * so an insert pushes widths off the end and a delete pulls the last
 * one along at it. An inserted column takes the width of the one it
 * pushed aside, which is the only answer that does not make a
 * carefully-widened column suddenly narrow while its neighbour is
 * wide. It is the same arithmetic the cells get, on an array instead
 * of a map.
 */
function shiftWidths(widths: readonly number[], shift: Shift): number[] {
  if (shift.axis !== 'column') {
    return [...widths];
  }
  const next = [...widths];
  const removed = -shift.by;
  if (shift.by > 0) {
    for (let column = widths.length - 1; column >= shift.at; column--) {
      const from = column - shift.by;
      next[column] = from >= shift.at ? widths[from] : widths[shift.at];
    }
  } else {
    for (let column = shift.at; column < widths.length; column++) {
      next[column] = widths[column + removed] ?? widths[widths.length - 1];
    }
  }
  return next;
}

/**
 * A set of row indices, moved by a shift.
 *
 * Hiding is by index, so an insert above a hidden row must not reveal
 * it and hide its neighbour instead.
 */
function shiftRows(rows: Set<number>, held: readonly number[], shift: Shift): void {
  rows.clear();
  for (const row of held) {
    const moved = shiftIndex(row, shift);
    if (moved !== -1) {
      rows.add(moved);
    }
  }
}
