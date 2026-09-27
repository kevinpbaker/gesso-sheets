import { columnIndex, inBounds, MAX_COLUMNS, parseRef, type RangeRef } from './A1';
import type { Ast } from './Ast';
import { isSheetFunction } from './Functions';
import { FormulaSyntaxError, parseFormula } from './Parser';
import { shiftFormula, shiftRange, type Shift } from './Shift';

/**
 * Names for ranges: `Sales` instead of `B2:B97`.
 *
 * The point is not brevity, it is that a formula stops needing a
 * decoder. `=SUM(Sales)/COUNT(Sales)` says what it works out;
 * `=SUM(B2:B97)/COUNT(B2:B97)` says where it looks, and the reader
 * has to go and find out what lives there.
 *
 * ## What a name may be
 *
 * The rules are Excel's, and every one of them exists to stop a name
 * being mistaken for something else.
 *
 * A name **may not look like a cell reference**. `A1` names a cell
 * and always will; a sheet where it could also name a range is a
 * sheet where the same formula means two things depending on a table
 * somebody cannot see. `AB12` is out for the same reason, and so is
 * a bare column letter, because `A:A` is a reference too.
 *
 * A name **starts with a letter or an underscore** and continues with
 * letters, digits, underscores and full stops. Starting with a digit
 * would collide with numbers; a space would collide with everything.
 *
 * Names are **case-insensitive**, because `sales` and `Sales` being
 * different ranges is a bug nobody would ever find. The case somebody
 * typed is kept for showing back to them.
 */

export interface NamedRange {
  /** As it was typed, for showing. Lookup is case-insensitive. */
  readonly name: string;
  readonly range: RangeRef;
}

/**
 * A name that holds a formula rather than a range.
 *
 * `TaxRate` as `=0.2`, `Gross` as `=Net*(1+TaxRate)` — a calculation
 * with a name — and, the reason this exists, `Double` as
 * `=LAMBDA(x, x*2)`: a function, called by its name as `=Double(A1)`
 * like any other. The formula is kept as written, `=` and all, and
 * parsed once when it is defined.
 */
export interface NamedFormula {
  readonly name: string;
  readonly formula: string;
  readonly ast: Ast;
}

/** Everything a name can hold. */
export type DefinedName = NamedRange | NamedFormula;

/** A name as a file or an undo step keeps it: a formula without its tree. */
export type StoredDefinedName = NamedRange | { readonly name: string; readonly formula: string };

export function isNamedRange(entry: DefinedName | StoredDefinedName): entry is NamedRange {
  return 'range' in entry;
}

/** Why a name was refused, in words somebody can act on. */
export type NameProblem = 'empty' | 'shape' | 'reference' | 'long' | 'function' | 'formula';

const SHAPE = /^[A-Za-z_][A-Za-z0-9_.]*$/;
/** Excel's limit, and long enough that nothing real reaches it. */
const MAX_LENGTH = 255;

/**
 * Whether a name may be used, and why not when it may not.
 *
 * Null for a name that is fine, which reads as "no problem".
 */
export function nameProblem(name: string): NameProblem | null {
  const trimmed = name.trim();
  if (trimmed === '') {
    return 'empty';
  }
  if (trimmed.length > MAX_LENGTH) {
    return 'long';
  }
  /**
   * Before the shape test, so `$A$1` is refused for being a
   * reference rather than for its dollar signs. Both refuse it; only
   * one of them tells somebody anything.
   */
  if (parseRef(trimmed) !== null || isColumnName(trimmed)) {
    return 'reference';
  }
  /**
   * A function's name is taken.
   *
   * The evaluator resolves functions before names, so a range called
   * `TODAY` could be defined and never read — the worst kind of
   * refusal, which is no refusal and no effect. Saying so is the
   * honest version.
   */
  if (isSheetFunction(trimmed.toUpperCase())) {
    return 'function';
  }
  if (!SHAPE.test(trimmed)) {
    return 'shape';
  }
  return null;
}

/**
 * Whether a word is a column of *this* sheet.
 *
 * `columnIndex` is bijective base-26 with no upper bound, so it reads
 * `SALES` as a perfectly good column number — three hundred million
 * or so. The sheet stops at XFD, and a word longer than three letters
 * is a word, which is the bound that makes this the question it was
 * meant to be.
 */
function isColumnName(text: string): boolean {
  if (text.length > 3) {
    return false;
  }
  const index = columnIndex(text);
  return index !== null && index < MAX_COLUMNS;
}

/** The sentence shown when a name is refused. */
export function nameProblemText(problem: NameProblem): string {
  switch (problem) {
    case 'empty':
      return 'A name cannot be empty.';
    case 'long':
      return `A name cannot be longer than ${MAX_LENGTH} characters.`;
    case 'shape':
      return 'A name starts with a letter or underscore, and holds letters, digits, underscores and full stops.';
    case 'reference':
      return 'That is a cell reference, so it already means something else.';
    case 'function':
      return 'That is the name of a function, which the sheet reaches for first.';
    case 'formula':
      return 'The formula after the = does not parse, so the name would hold nothing.';
  }
}

/** A formula for a name, parsed; the leading `=` is optional. Null if it does not parse. */
function parsedFormula(formula: string): Ast | null {
  const text = formula.trim().replace(/^=/, '');
  if (text === '') {
    return null;
  }
  try {
    return parseFormula(text);
  } catch (error) {
    if (error instanceof FormulaSyntaxError) {
      return null;
    }
    throw error;
  }
}

/**
 * The names a sheet knows.
 *
 * Keyed by the upper-cased name, holding the case somebody typed
 * alongside — so a lookup cannot depend on capitalisation and a
 * listing can still show `Sales` rather than `SALES`.
 */
export class Names {
  private readonly byKey = new Map<string, DefinedName>();

  get size(): number {
    return this.byKey.size;
  }

  /** The range a name stands for, or null — also for a name that holds a formula. */
  rangeOf(name: string): RangeRef | null {
    const entry = this.byKey.get(name.toUpperCase());
    return entry !== undefined && isNamedRange(entry) ? entry.range : null;
  }

  /** The formula a name holds, parsed, or null — also for a name that holds a range. */
  formulaOf(name: string): Ast | null {
    const entry = this.byKey.get(name.toUpperCase());
    return entry !== undefined && !isNamedRange(entry) ? entry.ast : null;
  }

  /** Whatever a name holds, or undefined. */
  get(name: string): DefinedName | undefined {
    return this.byKey.get(name.toUpperCase());
  }

  has(name: string): boolean {
    return this.byKey.has(name.toUpperCase());
  }

  /**
   * Defines or redefines a name.
   *
   * Returns the problem when the name is not one, and does nothing;
   * the caller is expected to say so rather than to have checked
   * first, because there is exactly one place that should hold the
   * rules and this is it.
   */
  define(name: string, range: RangeRef): NameProblem | null {
    const problem = nameProblem(name);
    if (problem !== null) {
      return problem;
    }
    const trimmed = name.trim();
    this.byKey.set(trimmed.toUpperCase(), { name: trimmed, range });
    return null;
  }

  /**
   * Gives a name a formula to hold, or says why not: the name's own
   * rules, and then a formula that parses. `=LAMBDA(x, x*2)` makes a
   * function; anything else is a calculation with a name.
   */
  defineFormula(name: string, formula: string): NameProblem | null {
    const problem = nameProblem(name);
    if (problem !== null) {
      return problem;
    }
    const ast = parsedFormula(formula);
    if (ast === null) {
      return 'formula';
    }
    const trimmed = name.trim();
    const text = formula.trim();
    this.byKey.set(trimmed.toUpperCase(), { name: trimmed, formula: text.startsWith('=') ? text : `=${text}`, ast });
    return null;
  }

  remove(name: string): boolean {
    return this.byKey.delete(name.toUpperCase());
  }

  /** Every name, in the order somebody would read them. */
  all(): DefinedName[] {
    return [...this.byKey.values()].sort((a, b) => a.name.localeCompare(b.name));
  }

  /** The names that hold ranges, which are the ones a name box can go to. */
  ranges(): NamedRange[] {
    return this.all().filter(isNamedRange);
  }

  /**
   * Moves every name's range with an insert or a delete.
   *
   * A name is a reference by another spelling, so it shifts like one:
   * rows inserted above `Sales` move it down, and a name whose every
   * cell was deleted is removed rather than left pointing at nothing.
   * Leaving it would give `#REF!` from a formula that names something
   * that still appears to exist, which is worse than the name being
   * gone.
   */
  shift(shift: Shift): void {
    for (const [key, entry] of [...this.byKey]) {
      if (!isNamedRange(entry)) {
        // A formula's references move as a cell's would; its
        // unqualified ones belong to whichever sheet reads it, so only
        // the ones that name the shifted sheet move.
        const moved = shiftFormula(entry.formula, shift);
        if (moved !== entry.formula) {
          const ast = parsedFormula(moved);
          if (ast !== null) {
            this.byKey.set(key, { name: entry.name, formula: moved, ast });
          }
        }
        continue;
      }
      const moved = shiftRange(entry.range, shift);
      // `shiftRange` marks a range whose every cell was deleted by
      // putting its corners off the sheet, which is what makes a
      // formula print `#REF!`. A *name* has somewhere better to go:
      // away. A name that still appears to exist and answers `#REF!`
      // is worse than one that is gone.
      const alive =
        inBounds(moved.start.row, moved.start.column) && inBounds(moved.end.row, moved.end.column);
      if (alive) {
        this.byKey.set(key, { name: entry.name, range: moved });
      } else {
        this.byKey.delete(key);
      }
    }
  }

  /**
   * Replaces the whole table, for loading a file and for undo. A stored
   * formula is parsed again; one that no longer parses is left out
   * rather than kept as a name that holds nothing.
   */
  restore(entries: readonly (DefinedName | StoredDefinedName)[]): void {
    this.byKey.clear();
    for (const entry of entries) {
      if (isNamedRange(entry)) {
        this.byKey.set(entry.name.toUpperCase(), entry);
        continue;
      }
      const ast = 'ast' in entry ? entry.ast : parsedFormula(entry.formula);
      if (ast !== null) {
        this.byKey.set(entry.name.toUpperCase(), { name: entry.name, formula: entry.formula, ast });
      }
    }
  }
}
