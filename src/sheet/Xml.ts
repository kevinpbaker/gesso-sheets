/**
 * Enough XML to read a spreadsheet file, and no more.
 *
 * An `.xlsx` is a zip of XML parts, and the application worker — where
 * reading one belongs — has no `DOMParser`: that is a window's, and a
 * worker is not a window. So this is a small reader of its own, and
 * small on purpose. What OOXML uses is elements, attributes, text,
 * the five named entities, numeric character references, CDATA,
 * comments and a declaration; what it does not use — a DTD, entities
 * defined in one, processing instructions that mean anything — is
 * skipped rather than understood, and that is safe *because* nothing
 * here expands an entity it did not define. A file cannot make this
 * fetch anything or grow without bound.
 *
 * Namespace prefixes are dropped from names: `x:row` and `row` are the
 * same element to a reader of spreadsheets, and producers disagree
 * about which to write.
 */

export interface XmlElement {
  readonly name: string;
  readonly attributes: Readonly<Record<string, string>>;
  readonly children: readonly XmlElement[];
  /** The element's own text, concatenated, with entities resolved. */
  readonly text: string;
}

export class XmlSyntaxError extends Error {}

const NAMED: Readonly<Record<string, string>> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };

/** Entities resolved, as XML defines them and nothing more. */
export function unescapeXml(text: string): string {
  if (!text.includes('&')) {
    return text;
  }
  return text.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[a-zA-Z]+);/g, (whole, body: string) => {
    if (body[0] === '#') {
      const code = body[1] === 'x' ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10);
      return Number.isFinite(code) && code >= 0 && code <= 0x10ffff ? String.fromCodePoint(code) : whole;
    }
    return NAMED[body] ?? whole;
  });
}

function localName(qualified: string): string {
  const colon = qualified.indexOf(':');
  return colon === -1 ? qualified : qualified.slice(colon + 1);
}

const ATTRIBUTE = /([^\s=/>]+)\s*=\s*("([^"]*)"|'([^']*)')/g;

/**
 * Reads a document and answers its root element.
 *
 * One pass over the string with an index, rather than a regular
 * expression over the whole of it, so a worksheet of a million cells
 * is read in time proportional to its length and a malformed one
 * fails at the place it is malformed.
 */
export function parseXml(source: string): XmlElement {
  interface Open {
    name: string;
    attributes: Record<string, string>;
    children: XmlElement[];
    text: string;
  }
  const stack: Open[] = [];
  let root: XmlElement | null = null;
  let at = 0;

  const close = (): void => {
    const done = stack.pop()!;
    const element: XmlElement = done;
    if (stack.length === 0) {
      root = element;
    } else {
      stack[stack.length - 1].children.push(element);
    }
  };

  while (at < source.length) {
    const lt = source.indexOf('<', at);
    const textEnd = lt === -1 ? source.length : lt;
    if (textEnd > at && stack.length > 0) {
      stack[stack.length - 1].text += unescapeXml(source.slice(at, textEnd));
    }
    if (lt === -1) {
      break;
    }
    if (source.startsWith('<!--', lt)) {
      const end = source.indexOf('-->', lt + 4);
      at = end === -1 ? source.length : end + 3;
      continue;
    }
    if (source.startsWith('<![CDATA[', lt)) {
      const end = source.indexOf(']]>', lt + 9);
      if (end === -1) {
        throw new XmlSyntaxError('A CDATA section is never closed.');
      }
      if (stack.length > 0) {
        stack[stack.length - 1].text += source.slice(lt + 9, end);
      }
      at = end + 3;
      continue;
    }
    if (source[lt + 1] === '?' || source[lt + 1] === '!') {
      // A declaration, a processing instruction or a DOCTYPE: skipped,
      // DOCTYPE and all, which is what keeps an entity it defines from
      // meaning anything here.
      const end = source.indexOf('>', lt);
      at = end === -1 ? source.length : end + 1;
      continue;
    }
    const gt = tagEnd(source, lt);
    if (gt === -1) {
      throw new XmlSyntaxError('A tag is never closed.');
    }
    const tag = source.slice(lt + 1, gt);
    at = gt + 1;
    if (tag[0] === '/') {
      const name = localName(tag.slice(1).trim());
      if (stack.length === 0 || stack[stack.length - 1].name !== name) {
        throw new XmlSyntaxError(`</${name}> closes an element that is not open.`);
      }
      close();
      continue;
    }
    const selfClosing = tag.endsWith('/');
    const body = selfClosing ? tag.slice(0, -1) : tag;
    const space = body.search(/\s/);
    const name = localName(space === -1 ? body : body.slice(0, space));
    const attributes: Record<string, string> = {};
    if (space !== -1) {
      ATTRIBUTE.lastIndex = 0;
      const rest = body.slice(space);
      let match: RegExpExecArray | null;
      while ((match = ATTRIBUTE.exec(rest)) !== null) {
        attributes[localName(match[1])] = unescapeXml(match[3] ?? match[4] ?? '');
      }
    }
    stack.push({ name, attributes, children: [], text: '' });
    if (selfClosing) {
      close();
    }
  }
  if (stack.length > 0) {
    throw new XmlSyntaxError(`<${stack[stack.length - 1].name}> is never closed.`);
  }
  if (root === null) {
    throw new XmlSyntaxError('There is no element in this document.');
  }
  return root;
}

/**
 * Where a tag ends: the first `>` outside quotes.
 *
 * Not the first `>`, because XML lets one stand unescaped inside an
 * attribute value and a spreadsheet's number formats use it —
 * `[>=100]0.0` is an ordinary `formatCode`.
 */
function tagEnd(source: string, from: number): number {
  let quote = '';
  for (let at = from + 1; at < source.length; at++) {
    const character = source[at];
    if (quote !== '') {
      if (character === quote) {
        quote = '';
      }
    } else if (character === '"' || character === "'") {
      quote = character;
    } else if (character === '>') {
      return at;
    }
  }
  return -1;
}

/** The first child of that name, or null. */
export function child(element: XmlElement | null | undefined, name: string): XmlElement | null {
  return element?.children.find(each => each.name === name) ?? null;
}

/** Every child of that name. */
export function children(element: XmlElement | null | undefined, name: string): XmlElement[] {
  return element === null || element === undefined ? [] : element.children.filter(each => each.name === name);
}
