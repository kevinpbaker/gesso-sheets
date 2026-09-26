/**
 * Bytes to and from base64, for a file crossing on a command.
 *
 * A module of its own because both workers need it and the render
 * worker must not load the engine: the reader it would otherwise sit
 * beside imports the parser, and that is exactly the code the render
 * worker's bundle is kept free of.
 */

/** Bytes from the base64 a command carried them in. */
export function bytesOfBase64(text: string): Uint8Array<ArrayBuffer> {
  const binary = atob(text);
  const bytes = new Uint8Array(binary.length);
  for (let at = 0; at < binary.length; at++) {
    bytes[at] = binary.charCodeAt(at);
  }
  return bytes;
}

/**
 * Bytes as base64, for a command to carry.
 *
 * A command's arguments cross as plain data, and an `ArrayBuffer` is
 * not — `requirePlainData` refuses anything with a prototype of its
 * own — so the file travels as a string a third larger than itself.
 * In chunks, because `String.fromCharCode(...bytes)` over a whole
 * workbook overflows the call stack.
 */
export function base64OfBytes(bytes: Uint8Array): string {
  let binary = '';
  for (let at = 0; at < bytes.length; at += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(at, at + 0x8000));
  }
  return btoa(binary);
}
