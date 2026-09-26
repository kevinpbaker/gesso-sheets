import { afterEach, describe, expect, it } from 'vitest';

import { OpfsSheetRepository, SheetReadError } from './OpfsSheetRepository';

/**
 * Reading the workbook back, when the file is not simply there.
 *
 * The case that lost data: a sync access handle is exclusive across
 * the origin, so a second tab asking while the first holds one is
 * refused, and a refused read used to come back as "nothing stored" —
 * which the service takes as leave to save a fresh, empty workbook
 * over the real one. The OPFS here is a double with exactly that
 * behaviour: it refuses the handle while "another tab" has it.
 */

class FakeHandle {
  constructor(private readonly file: FakeFile) {}
  getSize(): number {
    return this.file.bytes.length;
  }
  read(buffer: Uint8Array): number {
    buffer.set(this.file.bytes);
    return this.file.bytes.length;
  }
  write(buffer: Uint8Array): number {
    this.file.bytes = new Uint8Array(buffer);
    return buffer.length;
  }
  truncate(): void {
    this.file.bytes = new Uint8Array(0);
  }
  flush(): void {}
  close(): void {}
}

class FakeFile {
  bytes = new Uint8Array(0);
  /** How many more times a handle is refused, as while another tab holds one. */
  heldFor = 0;
  createSyncAccessHandle(): Promise<FakeHandle> {
    if (this.heldFor > 0) {
      this.heldFor--;
      const refused = new Error('Access Handles cannot be created if there is another open Access Handle.');
      refused.name = 'NoModificationAllowedError';
      return Promise.reject(refused);
    }
    return Promise.resolve(new FakeHandle(this));
  }
}

const file = new FakeFile();
const scope = globalThis as unknown as { navigator?: unknown };
const before = scope.navigator;

function install(): void {
  Object.defineProperty(globalThis, 'navigator', {
    configurable: true,
    value: { storage: { getDirectory: () => Promise.resolve({ getFileHandle: () => Promise.resolve(file) }) } }
  });
}

afterEach(() => {
  Object.defineProperty(globalThis, 'navigator', { configurable: true, value: before });
  file.bytes = new Uint8Array(0);
  file.heldFor = 0;
});

const WORKBOOK = JSON.stringify({
  version: 3,
  active: 0,
  names: [],
  sheets: [
    {
      name: 'Sheet1',
      colour: null,
      cells: [{ row: 0, column: 0, input: 'precious' }],
      palette: [],
      formats: [],
      regions: { sheet: 0, rows: [], columns: [] },
      merges: [],
      conditional: [],
      validations: [],
      charts: [],
      frozenRows: 0,
      frozenColumns: 0,
      hiddenRows: [],
      columnWidths: []
    }
  ]
});

describe('reading the workbook back', () => {
  it('answers null for a file with nothing in it, which is leave to start fresh', async () => {
    install();
    expect(await new OpfsSheetRepository('a.json', 10).load()).toBeNull();
  });

  it('waits for a file another tab is holding, and reads it', async () => {
    install();
    file.bytes = new TextEncoder().encode(WORKBOOK);
    file.heldFor = 3;
    const snapshot = await new OpfsSheetRepository('a.json', 10).load();
    expect(snapshot?.sheets[0].cells).toContainEqual({ row: 0, column: 0, input: 'precious' });
  });

  /** Never null: null would be saved over, and the file is there. */
  it('refuses to say a file is empty when it could not read it', async () => {
    install();
    file.bytes = new TextEncoder().encode(WORKBOOK);
    file.heldFor = 1_000;
    await expect(new OpfsSheetRepository('a.json', 10).load()).rejects.toThrow(SheetReadError);
  });

  it('refuses a file it cannot parse, rather than calling it empty', async () => {
    install();
    file.bytes = new TextEncoder().encode('{"not": "a workbook"');
    await expect(new OpfsSheetRepository('a.json', 10).load()).rejects.toThrow('cannot read');
  });
});
