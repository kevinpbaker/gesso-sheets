import { Sheet } from '../sheet/Sheet';

/** One cell's text, before and after. */
interface Edit {
  readonly row: number;
  readonly column: number;
  readonly before: string;
  readonly after: string;
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
    this.sheet.setCell(row, column, input);
    this.record({ row, column, before, after: input });
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
      this.sheet.setCell(edit.row, edit.column, edit.before);
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
      this.sheet.setCell(edit.row, edit.column, edit.after);
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

  /** What the formula bar shows for the active cell. */
  get activeInput(): string {
    return this.sheet.input(this.selection.row, this.selection.column);
  }
}
