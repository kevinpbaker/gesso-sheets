import type { Block } from './Tsv';

/**
 * Comma-separated values, read and written.
 *
 * **Not TSV with a different separator**, which is why this is not a
 * parameter on `Tsv.ts`. The two share the quoting rule and almost
 * nothing else that matters: a clipboard block is what this sheet or
 * another spreadsheet put there a moment ago, and a CSV is a file
 * that has been through an export dialog, a mail client and a text
 * editor on its way here. So a CSV can open with a byte-order mark,
 * say which separator it uses on a line of its own, use a semicolon
 * because its author's locale writes a comma where this one writes a
 * point, and end in whatever line ending the last program to touch it
 * preferred.
 *
 * What it does not decide is what a cell *means*. `=1+2` arriving in
 * a CSV is the four characters `=1+2`, and whether they are written
 * into the sheet as a formula is the importer's question; see
 * `SheetService.importCsv`, which answers it with no.
 */

export type CsvDelimiter = ',' | ';' | '\t';

export interface ParsedCsv {
  readonly rows: Block;
  /** Which separator the file turned out to use. */
  readonly delimiter: CsvDelimiter;
}

const DELIMITERS: readonly CsvDelimiter[] = [',', ';', '\t'];

/**
 * Reads a CSV.
 *
 * The separator is the file's own `sep=` line when it has one — Excel
 * writes it, and it is the only place a CSV ever says what it is —
 * and otherwise whichever of comma, semicolon and tab occurs most
 * often in the first record *outside quotes*. Outside quotes because
 * a first row of `"Smith, J";"Jones, K"` has more commas than
 * semicolons and is plainly semicolon-separated. A tie goes to the
 * comma, which is what the C stands for.
 *
 * Rows are padded to the widest, as a pasted block's are, and one
 * trailing line ending is dropped rather than read as a row of
 * blanks. A blank line in the middle is kept: it is a row somebody
 * left empty, and closing it up would move every row below it.
 */
export function parseCsv(text: string): ParsedCsv {
  let body = text.startsWith('﻿') ? text.slice(1) : text;

  let delimiter: CsvDelimiter | null = null;
  const hint = /^sep=(.)(\r\n|\n|\r|$)/.exec(body);
  if (hint !== null && (DELIMITERS as readonly string[]).includes(hint[1])) {
    delimiter = hint[1] as CsvDelimiter;
    body = body.slice(hint[0].length);
  }
  delimiter ??= sniff(body);

  const rows: string[][] = [];
  let row: string[] = [];
  let cell = '';
  let quoted = false;
  let at = 0;

  const endCell = (): void => {
    row.push(cell);
    cell = '';
  };
  const endRow = (): void => {
    endCell();
    rows.push(row);
    row = [];
  };

  while (at < body.length) {
    const character = body[at];

    if (quoted) {
      if (character === '"') {
        if (body[at + 1] === '"') {
          cell += '"';
          at += 2;
          continue;
        }
        quoted = false;
        at++;
        continue;
      }
      // A line ending inside quotes is part of the cell, and it is
      // normalised on the way in: a cell is not a file, and one that
      // held `\r\n` would draw a stray glyph and compare unequal to
      // the same text typed by hand.
      if (character === '\r') {
        cell += '\n';
        at += body[at + 1] === '\n' ? 2 : 1;
        continue;
      }
      cell += character;
      at++;
      continue;
    }

    if (character === '"' && cell === '') {
      quoted = true;
      at++;
      continue;
    }
    if (character === delimiter) {
      endCell();
      at++;
      continue;
    }
    if (character === '\r' || character === '\n') {
      endRow();
      at += character === '\r' && body[at + 1] === '\n' ? 2 : 1;
      continue;
    }
    cell += character;
    at++;
  }

  if (cell !== '' || row.length > 0 || quoted) {
    endRow();
  }
  return { rows: rectangular(rows), delimiter };
}

/** The separator the first record uses most, counted outside quotes. */
function sniff(body: string): CsvDelimiter {
  const counts = new Map<CsvDelimiter, number>(DELIMITERS.map(each => [each, 0]));
  let quoted = false;
  for (let at = 0; at < body.length; at++) {
    const character = body[at];
    if (character === '"') {
      quoted = !quoted;
      continue;
    }
    if (quoted) {
      continue;
    }
    if (character === '\r' || character === '\n') {
      break;
    }
    const seen = counts.get(character as CsvDelimiter);
    if (seen !== undefined) {
      counts.set(character as CsvDelimiter, seen + 1);
    }
  }
  let best: CsvDelimiter = ',';
  for (const each of DELIMITERS) {
    if ((counts.get(each) ?? 0) > (counts.get(best) ?? 0)) {
      best = each;
    }
  }
  return best;
}

function rectangular(rows: string[][]): Block {
  const width = rows.reduce((widest, row) => Math.max(widest, row.length), 0);
  for (const row of rows) {
    while (row.length < width) {
      row.push('');
    }
  }
  return rows;
}

/**
 * Writes a CSV, the way RFC 4180 and Excel both read one.
 *
 * `\r\n` between records and after the last, because that is the
 * RFC's line ending and the one every reader accepts. A cell is quoted
 * when it holds the separator, a quote or a line ending — and also
 * when it starts or ends with a space, which several readers trim
 * from an unquoted field and which is part of what somebody typed.
 */
export function toCsv(rows: Block, delimiter: CsvDelimiter = ','): string {
  if (rows.length === 0) {
    return '';
  }
  return rows.map(row => row.map(cell => quote(cell, delimiter)).join(delimiter)).join('\r\n') + '\r\n';
}

function quote(cell: string, delimiter: CsvDelimiter): string {
  const needs =
    cell.includes(delimiter) ||
    /["\r\n]/.test(cell) ||
    (cell.length > 0 && (cell[0] === ' ' || cell[cell.length - 1] === ' '));
  return needs ? `"${cell.replace(/"/g, '""')}"` : cell;
}

/**
 * Whether a text cell would be read as a formula by whatever opens
 * the file next.
 *
 * Spreadsheet injection is a CSV with `=HYPERLINK(...)` in a name
 * field, and the defence OWASP describes is the one used here: a
 * leading apostrophe on any *text* value that starts with one of the
 * characters that begin a formula somewhere. Only text — a negative
 * number starts with a minus and is a number, and prefixing it would
 * turn every loss in a ledger into a string.
 */
export function looksLikeFormula(text: string): boolean {
  return /^[=+\-@\t\r]/.test(text);
}
