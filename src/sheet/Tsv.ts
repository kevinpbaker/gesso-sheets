/**
 * The clipboard's shape for a block of cells.
 *
 * Tab-separated rows is what every spreadsheet reads and writes, and
 * the reason this phase's exit criterion is a paste out of Excel: the
 * format is only worth having if the real one lands.
 *
 * What makes it more than `split('\t')` is quoting. A cell holding a
 * tab, a newline or a quote is written in double quotes with its own
 * quotes doubled, exactly as CSV does — and a sheet that wrote them
 * raw would turn one cell into two on the way back.
 */

/** A block of cells, row-major. */
export type Block = readonly (readonly string[])[];

export function toTsv(block: Block): string {
  return block.map(row => row.map(quote).join('\t')).join('\n');
}

function quote(cell: string): string {
  if (!/[\t\n\r"]/.test(cell)) {
    return cell;
  }
  return `"${cell.replace(/"/g, '""')}"`;
}

/**
 * Reads a block back.
 *
 * Rows are padded to the width of the widest, because a block has to
 * be a rectangle for anything downstream to reason about it, and a
 * ragged paste out of a text editor is a thing people do.
 *
 * A trailing newline is dropped: Excel puts one on the end of what it
 * copies, and reading it as an extra row of blanks would wipe a row
 * of the sheet the person did not mean to touch.
 */
export function fromTsv(text: string): Block {
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

  while (at < text.length) {
    const character = text[at];

    if (quoted) {
      if (character === '"') {
        if (text[at + 1] === '"') {
          cell += '"';
          at += 2;
          continue;
        }
        quoted = false;
        at++;
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
    if (character === '\t') {
      endCell();
      at++;
      continue;
    }
    if (character === '\r' || character === '\n') {
      endRow();
      at += character === '\r' && text[at + 1] === '\n' ? 2 : 1;
      continue;
    }
    cell += character;
    at++;
  }

  if (cell !== '' || row.length > 0) {
    endRow();
  }
  return rectangular(rows);
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

/** Whether a string is worth reading as a block rather than as one cell. */
export function looksTabular(text: string): boolean {
  return text.includes('\t') || text.includes('\n') || text.includes('\r');
}
