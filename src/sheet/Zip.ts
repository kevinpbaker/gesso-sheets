/**
 * The entries of a zip file, read from its central directory.
 *
 * An `.xlsx` is a zip, and this is what it takes to open one: find the
 * end record, walk the directory it points at, and read each entry
 * from where its local header says the bytes start. Stored entries are
 * the bytes; deflated ones go through `inflate`, which is handed in —
 * `DecompressionStream` in a browser or in node, both of which have it
 * — so that nothing here reaches for a platform, which is the rule
 * this directory is under.
 *
 * What is refused, by name: Zip64 (a spreadsheet this application can
 * hold is nowhere near four gigabytes, and one that needs Zip64 is not
 * one it could open anyway), encryption, and any method but stored and
 * deflate. Those are the whole of the formats' corners that a
 * spreadsheet ever lands in, and refusing them with a sentence is
 * better than reading them wrongly.
 */

export type Inflate = (bytes: Uint8Array) => Promise<Uint8Array>;

export class ZipError extends Error {}

export interface ZipEntry {
  readonly name: string;
  readonly method: number;
  readonly compressedSize: number;
  readonly size: number;
  /** Where the entry's local header starts. */
  readonly offset: number;
  readonly flags: number;
}

const END_OF_DIRECTORY = 0x06054b50;
const DIRECTORY_ENTRY = 0x02014b50;
const LOCAL_HEADER = 0x04034b50;

/** The directory: every entry's name and where to find it. */
export function zipEntries(bytes: Uint8Array): ZipEntry[] {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  // The end record is 22 bytes plus a comment of up to 65,535, so it
  // is searched for backwards from the end over at most that much.
  let end = -1;
  for (let at = bytes.length - 22; at >= Math.max(0, bytes.length - 22 - 0xffff); at--) {
    if (view.getUint32(at, true) === END_OF_DIRECTORY) {
      end = at;
      break;
    }
  }
  if (end === -1) {
    throw new ZipError('This is not a zip file.');
  }
  const count = view.getUint16(end + 10, true);
  const directory = view.getUint32(end + 16, true);
  if (count === 0xffff || directory === 0xffffffff) {
    throw new ZipError('This file needs Zip64, which is not read here.');
  }
  const decoder = new TextDecoder();
  const entries: ZipEntry[] = [];
  let at = directory;
  for (let index = 0; index < count; index++) {
    if (at + 46 > bytes.length || view.getUint32(at, true) !== DIRECTORY_ENTRY) {
      throw new ZipError('The zip directory is damaged.');
    }
    const flags = view.getUint16(at + 8, true);
    const method = view.getUint16(at + 10, true);
    const compressedSize = view.getUint32(at + 20, true);
    const size = view.getUint32(at + 24, true);
    const nameLength = view.getUint16(at + 28, true);
    const extraLength = view.getUint16(at + 30, true);
    const commentLength = view.getUint16(at + 32, true);
    const offset = view.getUint32(at + 42, true);
    const name = decoder.decode(bytes.subarray(at + 46, at + 46 + nameLength));
    entries.push({ name, method, compressedSize, size, offset, flags });
    at += 46 + nameLength + extraLength + commentLength;
  }
  return entries;
}

/** One entry's bytes, decompressed. */
export async function zipRead(bytes: Uint8Array, entry: ZipEntry, inflate: Inflate): Promise<Uint8Array> {
  if ((entry.flags & 1) !== 0) {
    throw new ZipError(`${entry.name} is encrypted.`);
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (entry.offset + 30 > bytes.length || view.getUint32(entry.offset, true) !== LOCAL_HEADER) {
    throw new ZipError(`${entry.name} is damaged.`);
  }
  // The local header's own lengths, not the directory's: the two
  // extra fields are allowed to differ, and the data starts after the
  // local one.
  const nameLength = view.getUint16(entry.offset + 26, true);
  const extraLength = view.getUint16(entry.offset + 28, true);
  const start = entry.offset + 30 + nameLength + extraLength;
  const data = bytes.subarray(start, start + entry.compressedSize);
  if (entry.method === 0) {
    return data;
  }
  if (entry.method === 8) {
    return inflate(data);
  }
  throw new ZipError(`${entry.name} is compressed in a way this cannot read (method ${entry.method}).`);
}

/** Deflates raw bytes, without a zlib header — the method-8 body of a zip entry. */
export type Deflate = (bytes: Uint8Array) => Promise<Uint8Array>;

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc = CRC_TABLE[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

/**
 * A zip of the given files, in order.
 *
 * Deflated when a `deflate` is handed in — the platform's, like the
 * inflate the reader takes — and stored otherwise, which every reader
 * accepts and is only larger. A file that deflates to more than it was
 * is stored anyway. No zip64: a sheet this size is nowhere near 4GB.
 */
export async function zipWrite(
  files: readonly { readonly name: string; readonly bytes: Uint8Array }[],
  deflate?: Deflate
): Promise<Uint8Array> {
  const encoder = new TextEncoder();
  const locals: Uint8Array[] = [];
  const directory: Uint8Array[] = [];
  let offset = 0;
  for (const file of files) {
    const name = encoder.encode(file.name);
    const crc = crc32(file.bytes);
    const packed = deflate === undefined ? null : await deflate(file.bytes);
    const deflated = packed !== null && packed.length < file.bytes.length;
    const body = deflated ? packed : file.bytes;
    const method = deflated ? 8 : 0;

    const local = new Uint8Array(30 + name.length);
    const header = new DataView(local.buffer);
    header.setUint32(0, LOCAL_HEADER, true);
    header.setUint16(4, 20, true);
    // Bit 11: the name is UTF-8.
    header.setUint16(6, 0x0800, true);
    header.setUint16(8, method, true);
    header.setUint16(10, 0, true);
    header.setUint16(12, 0x21, true);
    header.setUint32(14, crc, true);
    header.setUint32(18, body.length, true);
    header.setUint32(22, file.bytes.length, true);
    header.setUint16(26, name.length, true);
    header.setUint16(28, 0, true);
    local.set(name, 30);
    locals.push(local, body);

    const entry = new Uint8Array(46 + name.length);
    const view = new DataView(entry.buffer);
    view.setUint32(0, DIRECTORY_ENTRY, true);
    view.setUint16(4, 20, true);
    view.setUint16(6, 20, true);
    view.setUint16(8, 0x0800, true);
    view.setUint16(10, method, true);
    view.setUint16(12, 0, true);
    view.setUint16(14, 0x21, true);
    view.setUint32(16, crc, true);
    view.setUint32(20, body.length, true);
    view.setUint32(24, file.bytes.length, true);
    view.setUint16(28, name.length, true);
    view.setUint32(42, offset, true);
    entry.set(name, 46);
    directory.push(entry);

    offset += local.length + body.length;
  }
  const directorySize = directory.reduce((sum, entry) => sum + entry.length, 0);
  const end = new Uint8Array(22);
  const tail = new DataView(end.buffer);
  tail.setUint32(0, END_OF_DIRECTORY, true);
  tail.setUint16(8, files.length, true);
  tail.setUint16(10, files.length, true);
  tail.setUint32(12, directorySize, true);
  tail.setUint32(16, offset, true);

  const out = new Uint8Array(offset + directorySize + end.length);
  let at = 0;
  for (const part of [...locals, ...directory, end]) {
    out.set(part, at);
    at += part.length;
  }
  return out;
}
