import { Sheet } from '../sheet/Sheet';

/** One cell's text, before and after. */
interface Edit {
  readonly row: number;
  readonly column: number;
  readonly before: string;
  readonly after: string;
}

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

  private readonly undoStack: Edit[] = [];
  private readonly redoStack: Edit[] = [];

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
    this.undoStack.push({ row, column, before, after: input });
    // A new edit is a new future; whatever was undone is unreachable.
    this.redoStack.length = 0;
  }

  undo(): boolean {
    const edit = this.undoStack.pop();
    if (edit === undefined) {
      return false;
    }
    this.sheet.setCell(edit.row, edit.column, edit.before);
    this.redoStack.push(edit);
    this.selection = { row: edit.row, column: edit.column, anchorRow: edit.row, anchorColumn: edit.column };
    return true;
  }

  redo(): boolean {
    const edit = this.redoStack.pop();
    if (edit === undefined) {
      return false;
    }
    this.sheet.setCell(edit.row, edit.column, edit.after);
    this.undoStack.push(edit);
    this.selection = { row: edit.row, column: edit.column, anchorRow: edit.row, anchorColumn: edit.column };
    return true;
  }

  setSelection(row: number, column: number, anchorRow: number, anchorColumn: number): void {
    this.selection = { row, column, anchorRow, anchorColumn };
  }

  /** What the formula bar shows for the active cell. */
  get activeInput(): string {
    return this.sheet.input(this.selection.row, this.selection.column);
  }
}
