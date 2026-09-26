import { describe, expect, it } from 'vitest';

import { zipEntries, zipRead, ZipError, type Inflate } from './Zip';
import { parseXml, XmlSyntaxError, child, children } from './Xml';

/**
 * The zip reader and the XML reader, which are what an `.xlsx` is
 * opened with.
 *
 * The zips are built here, by a writer small enough to read: stored
 * and deflated entries, a comment after the end record, the shapes a
 * producer is allowed to write. Real files from LibreOffice are read in
 * `Xlsx.spec.ts`; this is the format's edges, one at a time.
 */

const inflate: Inflate = async bytes => {
  const stream = new Blob([bytes as Uint8Array<ArrayBuffer>]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
};

async function deflate(bytes: Uint8Array): Promise<Uint8Array> {
  const stream = new Blob([bytes as Uint8Array<ArrayBuffer>]).stream().pipeThrough(new CompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

interface Part {
  name: string;
  text: string;
  method?: 0 | 8;
  flags?: number;
}

/** A zip, written the way the format says to. */
async function zip(parts: readonly Part[], comment = ''): Promise<Uint8Array> {
  const encoder = new TextEncoder();
  const chunks: Uint8Array[] = [];
  const directory: Uint8Array[] = [];
  let offset = 0;
  for (const part of parts) {
    const name = encoder.encode(part.name);
    const raw = encoder.encode(part.text);
    const method = part.method ?? 8;
    const data = method === 8 ? await deflate(raw) : raw;
    const local = new DataView(new ArrayBuffer(30));
    local.setUint32(0, 0x04034b50, true);
    local.setUint16(6, part.flags ?? 0, true);
    local.setUint16(8, method, true);
    local.setUint32(18, data.length, true);
    local.setUint32(22, raw.length, true);
    local.setUint16(26, name.length, true);
    chunks.push(new Uint8Array(local.buffer), name, data);
    const entry = new DataView(new ArrayBuffer(46));
    entry.setUint32(0, 0x02014b50, true);
    entry.setUint16(8, part.flags ?? 0, true);
    entry.setUint16(10, method, true);
    entry.setUint32(20, data.length, true);
    entry.setUint32(24, raw.length, true);
    entry.setUint16(28, name.length, true);
    entry.setUint32(42, offset, true);
    directory.push(new Uint8Array(entry.buffer), name);
    offset += 30 + name.length + data.length;
  }
  const directorySize = directory.reduce((sum, chunk) => sum + chunk.length, 0);
  const note = encoder.encode(comment);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true);
  end.setUint16(8, parts.length, true);
  end.setUint16(10, parts.length, true);
  end.setUint32(12, directorySize, true);
  end.setUint32(16, offset, true);
  end.setUint16(20, note.length, true);
  const all = [...chunks, ...directory, new Uint8Array(end.buffer), note];
  const out = new Uint8Array(all.reduce((sum, chunk) => sum + chunk.length, 0));
  let at = 0;
  for (const chunk of all) {
    out.set(chunk, at);
    at += chunk.length;
  }
  return out;
}

const text = (bytes: Uint8Array) => new TextDecoder().decode(bytes);

describe('reading a zip', () => {
  it('lists its entries and reads stored and deflated ones alike', async () => {
    const file = await zip([
      { name: 'xl/workbook.xml', text: '<workbook/>' },
      { name: '[Content_Types].xml', text: '<Types/>', method: 0 }
    ]);
    const entries = zipEntries(file);
    expect(entries.map(entry => [entry.name, entry.method])).toEqual([
      ['xl/workbook.xml', 8],
      ['[Content_Types].xml', 0]
    ]);
    expect(text(await zipRead(file, entries[0], inflate))).toBe('<workbook/>');
    expect(text(await zipRead(file, entries[1], inflate))).toBe('<Types/>');
  });

  it('finds the directory behind a comment', async () => {
    const file = await zip([{ name: 'a.xml', text: 'x'.repeat(5000) }], 'written by something chatty');
    const [entry] = zipEntries(file);
    expect(text(await zipRead(file, entry, inflate))).toHaveLength(5000);
  });

  it('refuses what is not a zip, and what is encrypted, with a sentence', async () => {
    expect(() => zipEntries(new TextEncoder().encode('Region,Units\n'))).toThrow(ZipError);
    const locked = await zip([{ name: 'a.xml', text: 'x', flags: 1 }]);
    await expect(zipRead(locked, zipEntries(locked)[0], inflate)).rejects.toThrow('a.xml is encrypted.');
  });
});

describe('reading XML', () => {
  it('reads elements, attributes and text, prefixes dropped', () => {
    const root = parseXml(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<x:worksheet xmlns:x="ns"><x:row r="1"><x:c r="A1" t="s"><x:v>0</x:v></x:c></x:row></x:worksheet>'
    );
    expect(root.name).toBe('worksheet');
    const cell = child(child(root, 'row'), 'c')!;
    expect(cell.attributes).toEqual({ r: 'A1', t: 's' });
    expect(child(cell, 'v')!.text).toBe('0');
  });

  it('resolves the entities XML defines, and only those', () => {
    const root = parseXml('<t a="&quot;x&quot;">1 &lt; 2 &amp; &#233;&#x1F600; &unknown;</t>');
    expect(root.text).toBe('1 < 2 & é😀 &unknown;');
    expect(root.attributes.a).toBe('"x"');
  });

  it('keeps a > inside an attribute, where number formats put one', () => {
    const root = parseXml('<numFmts><numFmt numFmtId="164" formatCode="[>=100]0.0;0"/></numFmts>');
    expect(children(root, 'numFmt')[0].attributes.formatCode).toBe('[>=100]0.0;0');
  });

  it('reads CDATA and skips comments and a DOCTYPE, entities and all', () => {
    const root = parseXml('<!DOCTYPE t [<!ENTITY boom "BOOM">]><t><!-- no --><![CDATA[<raw> &amp;]]>&boom;</t>');
    expect(root.text).toBe('<raw> &amp;&boom;');
  });

  it('refuses a document whose tags do not close', () => {
    expect(() => parseXml('<a><b></a>')).toThrow(XmlSyntaxError);
    expect(() => parseXml('<a>')).toThrow(XmlSyntaxError);
  });
});
