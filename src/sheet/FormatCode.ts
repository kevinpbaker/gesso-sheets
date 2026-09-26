import { dateOfSerial, weekdayOf } from './Dates';
import { isError, numberOfText, VALUE, type CellError, type CellValue } from './Values';

/**
 * Excel's number format codes: `#,##0.00`, `[h]:mm:ss`, `0.0%;[Red]-0.0%`.
 *
 * The sheet's own formats are a handful of named kinds, which is what
 * a toolbar can offer; a *code* is the language Excel files and `TEXT`
 * speak, and this reads it. It is written against Apache POI's format
 * test files — about a thousand codes with the text Excel produced for
 * each — which `FormatCode.spec.ts` replays.
 *
 * What is here: up to four sections, and conditions choosing between
 * them; literals, quoted text, escapes, `_` padding (a space) and `*`
 * fills (nothing, since there is no width to fill); colours (dropped);
 * currency in brackets; `General`; the digit placeholders `0` `#` `?`
 * with thousands separators, scaling commas, percent and scientific;
 * fractions; and dates, times, elapsed time and AM/PM. What is not:
 * locale identifiers beyond the currency symbol, and the rarest corners
 * of fraction layout, where Excel's own behaviour is hard to pin down.
 */

export interface FormatOptions {
  /** The 1904 date system, which Excel for the Mac used by default. */
  readonly date1904?: boolean;
}

type Token =
  | { readonly kind: 'literal'; readonly text: string }
  | { readonly kind: 'digit'; readonly char: '0' | '#' | '?' }
  | { readonly kind: 'point' }
  | { readonly kind: 'comma' }
  | { readonly kind: 'percent' }
  | { readonly kind: 'exponent'; readonly sign: '+' | '-'; readonly letter: string }
  | { readonly kind: 'slash' }
  | { readonly kind: 'text' }
  | { readonly kind: 'general' }
  | { readonly kind: 'date'; readonly code: string }
  | { readonly kind: 'elapsed'; readonly unit: 'h' | 'm' | 's'; readonly width: number }
  | { readonly kind: 'ampm'; readonly code: string };

interface Condition {
  readonly op: '<' | '<=' | '>' | '>=' | '=' | '<>';
  readonly value: number;
}

interface Section {
  readonly tokens: readonly Token[];
  readonly condition: Condition | null;
  readonly isDate: boolean;
}

/** Formats a value with a code, as `TEXT` and a cell do. */
export function formatCode(value: CellValue, code: string, options: FormatOptions = {}): string | CellError {
  if (isError(value)) {
    return value;
  }
  const sections = splitSections(code).map(parseSection);
  if (typeof value === 'string' || typeof value === 'boolean') {
    const text = typeof value === 'boolean' ? (value ? 'TRUE' : 'FALSE') : value;
    // A number written as text is formatted as the number, as Excel's
    // TEXT does; anything else is text.
    if (typeof value === 'string') {
      const number = numberOfText(value);
      if (!isError(number)) {
        return formatCode(number, code, options);
      }
    }
    return formatText(text, sections);
  }
  const number = value ?? 0;
  const chosen = chooseSection(number, sections);
  if (chosen === null) {
    return formatGeneral(number);
  }
  const { section, dropSign } = chosen;
  const shown = dropSign ? Math.abs(number) : number;
  if (section.isDate) {
    return formatDate(shown, section.tokens, options);
  }
  return formatNumber(shown, section.tokens);
}

// ---------------------------------------------------------------------------
// Sections
// ---------------------------------------------------------------------------

function splitSections(code: string): string[] {
  const sections: string[] = [];
  let current = '';
  let quoted = false;
  let bracket = false;
  for (let at = 0; at < code.length; at++) {
    const character = code[at];
    if (quoted) {
      current += character;
      if (character === '"') {
        quoted = false;
      }
      continue;
    }
    if (character === '\\' || character === '_' || character === '*') {
      current += character + (code[at + 1] ?? '');
      at++;
      continue;
    }
    if (character === '"') {
      quoted = true;
    } else if (character === '[') {
      bracket = true;
    } else if (character === ']') {
      bracket = false;
    } else if (character === ';' && !bracket) {
      sections.push(current);
      current = '';
      continue;
    }
    current += character;
  }
  sections.push(current);
  return sections;
}

const CONDITION = /^(<=|>=|<>|<|>|=)\s*(-?\d+(?:\.\d+)?)$/;

function parseSection(text: string): Section {
  const tokens: Token[] = [];
  let condition: Condition | null = null;
  let at = 0;
  const lower = text.toLowerCase();
  while (at < text.length) {
    const character = text[at];
    if (character === '"') {
      const end = text.indexOf('"', at + 1);
      const stop = end === -1 ? text.length : end;
      tokens.push({ kind: 'literal', text: text.slice(at + 1, stop) });
      at = stop + 1;
      continue;
    }
    if (character === '\\') {
      tokens.push({ kind: 'literal', text: text[at + 1] ?? '' });
      at += 2;
      continue;
    }
    if (character === '_') {
      // Space the width of the next character; a space is as close as
      // text can come.
      tokens.push({ kind: 'literal', text: ' ' });
      at += 2;
      continue;
    }
    if (character === '*') {
      // Repeat to fill the cell, which text has no width to do.
      at += 2;
      continue;
    }
    if (character === '[') {
      const end = text.indexOf(']', at);
      const inner = text.slice(at + 1, end === -1 ? text.length : end);
      at = end === -1 ? text.length : end + 1;
      const matched = CONDITION.exec(inner.trim());
      if (matched !== null) {
        condition = { op: matched[1] as Condition['op'], value: Number(matched[2]) };
      } else if (inner.startsWith('$')) {
        const symbol = inner.slice(1).split('-')[0];
        if (symbol !== '') {
          tokens.push({ kind: 'literal', text: symbol });
        }
      } else if (/^(h+|m+|s+)$/i.test(inner)) {
        tokens.push({ kind: 'elapsed', unit: inner[0].toLowerCase() as 'h' | 'm' | 's', width: inner.length });
      }
      // Anything else is a colour or a locale, which change nothing here.
      continue;
    }
    if (lower.startsWith('general', at)) {
      tokens.push({ kind: 'general' });
      at += 7;
      continue;
    }
    if (lower.startsWith('am/pm', at)) {
      tokens.push({ kind: 'ampm', code: text.slice(at, at + 5) });
      at += 5;
      continue;
    }
    if (lower.startsWith('a/p', at)) {
      tokens.push({ kind: 'ampm', code: text.slice(at, at + 3) });
      at += 3;
      continue;
    }
    if (character === '0' || character === '#' || character === '?') {
      tokens.push({ kind: 'digit', char: character });
      at++;
      continue;
    }
    if (character === '.') {
      tokens.push({ kind: 'point' });
      at++;
      continue;
    }
    if (character === ',') {
      tokens.push({ kind: 'comma' });
      at++;
      continue;
    }
    if (character === '%') {
      tokens.push({ kind: 'percent' });
      at++;
      continue;
    }
    if (character === '/') {
      tokens.push({ kind: 'slash' });
      at++;
      continue;
    }
    if (character === '@') {
      tokens.push({ kind: 'text' });
      at++;
      continue;
    }
    if ((character === 'e' || character === 'E') && (text[at + 1] === '+' || text[at + 1] === '-')) {
      tokens.push({ kind: 'exponent', sign: text[at + 1] as '+' | '-', letter: character });
      at += 2;
      continue;
    }
    if (/[ymdhs]/i.test(character)) {
      let end = at + 1;
      while (end < text.length && text[end].toLowerCase() === character.toLowerCase()) {
        end++;
      }
      tokens.push({ kind: 'date', code: text.slice(at, end).toLowerCase() });
      at = end;
      continue;
    }
    tokens.push({ kind: 'literal', text: character });
    at++;
  }
  const isDate = tokens.some(token => token.kind === 'date' || token.kind === 'elapsed' || token.kind === 'ampm');
  return { tokens, condition, isDate };
}

function holds(condition: Condition, value: number): boolean {
  switch (condition.op) {
    case '<':
      return value < condition.value;
    case '<=':
      return value <= condition.value;
    case '>':
      return value > condition.value;
    case '>=':
      return value >= condition.value;
    case '=':
      return value === condition.value;
    case '<>':
      return value !== condition.value;
  }
}

/**
 * Which section formats a number, and whether it shows the sign.
 *
 * Without conditions: one section is everything; two are the positive
 * and zero, then the negative; three add zero on its own. A negative
 * number in a section other than the first loses its minus sign, since
 * that section is where the code says how negatives look. With
 * conditions, the first that holds, then a section without one; and
 * nothing at all — a number no section speaks for — is General.
 */
function chooseSection(value: number, sections: readonly Section[]): { section: Section; dropSign: boolean } | null {
  const numeric = sections.slice(0, 3);
  if (numeric.some(section => section.condition !== null)) {
    for (let at = 0; at < numeric.length; at++) {
      const section = numeric[at];
      if (section.condition === null ? at > 0 : holds(section.condition, value)) {
        return { section, dropSign: at > 0 && value < 0 && sections.length > 1 && section.condition === null };
      }
    }
    return null;
  }
  if (numeric.length === 1 || (numeric.length === 2 && value >= 0) || (numeric.length >= 3 && value > 0)) {
    return { section: numeric[0], dropSign: false };
  }
  if (value < 0) {
    return { section: numeric[1], dropSign: true };
  }
  return { section: numeric[2] ?? numeric[0], dropSign: false };
}

// ---------------------------------------------------------------------------
// Text
// ---------------------------------------------------------------------------

function formatText(text: string, sections: readonly Section[]): string {
  const section =
    sections.length >= 4
      ? sections[3]
      : sections.length === 1 && !sections[0].tokens.some(token => token.kind === 'digit' || token.kind === 'general')
        ? sections[0]
        : null;
  if (section === null) {
    return text;
  }
  return section.tokens.map(token => (token.kind === 'text' ? text : token.kind === 'literal' ? token.text : '')).join('');
}

// ---------------------------------------------------------------------------
// General
// ---------------------------------------------------------------------------

/**
 * Excel's General, as `TEXT` gives it: as many digits as fit in eleven
 * characters, sign aside, and scientific with six significant digits
 * from 1E+11 up and below 1E-9.
 */
export function formatGeneral(value: number): string {
  if (value === 0) {
    return '0';
  }
  const sign = value < 0 ? '-' : '';
  const magnitude = Math.abs(value);
  if (magnitude < 1e11 && magnitude >= 1e-9) {
    const whole = magnitude >= 1 ? Math.floor(Math.log10(magnitude)) + 1 : 1;
    const text = magnitude.toFixed(Math.max(0, 10 - whole));
    const trimmed = text.includes('.') ? text.replace(/\.?0+$/, '') : text;
    if (trimmed.length <= 11 && trimmed !== '0') {
      return sign + trimmed;
    }
  }
  const [mantissa, exponent] = magnitude.toExponential(5).split('e');
  const power = Number(exponent);
  return `${sign}${mantissa.replace(/\.?0+$/, '')}E${power < 0 ? '-' : '+'}${String(Math.abs(power)).padStart(2, '0')}`;
}

// ---------------------------------------------------------------------------
// Numbers
// ---------------------------------------------------------------------------

function formatNumber(value: number, tokens: readonly Token[]): string {
  const negative = value < 0;
  let magnitude = Math.abs(value);
  const sign = negative ? '-' : '';

  if (tokens.some(token => token.kind === 'general')) {
    return sign + tokens.map(token => (token.kind === 'general' ? formatGeneral(magnitude) : token.kind === 'literal' ? token.text : '')).join('');
  }
  if (!tokens.some(token => token.kind === 'digit')) {
    // A section of literals only: "Zero", or nothing at all for a
    // code like `0;;` that hides negatives.
    return sign + tokens.map(token => (token.kind === 'literal' ? token.text : token.kind === 'percent' ? '%' : '')).join('');
  }

  // Scientific notation takes its own scale: percent and scaling commas
  // are printed but multiply nothing.
  const scientific = tokens.some(token => token.kind === 'exponent');
  for (const token of tokens) {
    if (token.kind === 'percent' && !scientific) {
      magnitude *= 100;
    }
  }

  const slash = tokens.findIndex(token => token.kind === 'slash');
  if (slash !== -1 && tokens.slice(slash + 1).some(token => token.kind === 'digit')) {
    return sign + formatFraction(magnitude, tokens, slash);
  }

  const exponentAt = tokens.findIndex(token => token.kind === 'exponent');
  const mantissaTokens = exponentAt === -1 ? tokens : tokens.slice(0, exponentAt);
  const pointAt = mantissaTokens.findIndex(token => token.kind === 'point');
  const integerTokens = pointAt === -1 ? mantissaTokens : mantissaTokens.slice(0, pointAt);
  const decimalTokens = pointAt === -1 ? [] : mantissaTokens.slice(pointAt + 1);

  // Commas: between integer placeholders a thousands separator; after
  // the last integer placeholder, each one divides by a thousand; before
  // any placeholder, a literal comma.
  let thousands = false;
  let scale = 0;
  const lastDigit = lastIndexOf(mantissaTokens, token => token.kind === 'digit');
  const firstDigit = mantissaTokens.findIndex(token => token.kind === 'digit');
  const commaRole = new Map<number, 'thousands' | 'scale' | 'literal'>();
  mantissaTokens.forEach((token, at) => {
    if (token.kind !== 'comma') {
      return;
    }
    if (at < firstDigit) {
      commaRole.set(at, 'literal');
    } else if (at > lastDigit || mantissaTokens[at + 1]?.kind === 'point') {
      commaRole.set(at, 'scale');
      scale++;
    } else if (pointAt === -1 || at < pointAt) {
      commaRole.set(at, 'thousands');
      thousands = true;
    } else {
      commaRole.set(at, 'literal');
    }
  });
  if (!scientific) {
    magnitude /= 1000 ** scale;
  }

  const decimals = decimalTokens.filter(token => token.kind === 'digit').length;
  let exponent = 0;
  let exponentText = '';
  if (exponentAt !== -1) {
    const exponentToken = tokens[exponentAt] as Extract<Token, { kind: 'exponent' }>;
    // The exponent is a multiple of the integer placeholders' count:
    // `0.0E+0` is scientific, `##0.0E+0` engineering.
    const step = Math.max(1, integerTokens.filter(token => token.kind === 'digit').length);
    if (magnitude !== 0) {
      exponent = Math.floor(Math.floor(Math.log10(magnitude)) / step) * step;
      magnitude /= 10 ** exponent;
      // Rounding can carry the mantissa to the next power.
      if (roundTo(magnitude, decimals) >= 10 ** step) {
        exponent += step;
        magnitude /= 10 ** step;
      }
    }
    const exponentDigits = tokens.slice(exponentAt + 1).filter(token => token.kind === 'digit');
    const zeros = exponentDigits.filter(token => token.char === '0').length;
    const spaces = exponentDigits.filter(token => token.char !== '#').length;
    const exponentSign = exponent < 0 ? '-' : exponentToken.sign === '+' ? '+' : '';
    const digits = exponentSign + String(Math.abs(exponent)).padStart(zeros, '0').padStart(spaces, ' ');
    // The sign goes with the digits, so a literal between the mark and
    // them sits before it: `e+|0|` is `e|+5|`.
    let rest = '';
    let used = false;
    for (const token of tokens.slice(exponentAt + 1)) {
      if (token.kind === 'digit') {
        if (!used) {
          rest += digits;
          used = true;
        }
      } else if (token.kind === 'literal') {
        rest += token.text;
      } else if (token.kind === 'percent') {
        rest += '%';
      }
    }
    exponentText = exponentToken.letter + (used ? rest : digits + rest);
  }

  const rounded = roundTo(magnitude, decimals);
  const [whole, fraction = ''] = rounded.toFixed(decimals).split('.');
  // Zero in scientific notation fills every integer placeholder.
  const places = integerTokens.filter(token => token.kind === 'digit').length;
  const integerDigits = whole !== '0' ? whole : scientific && rounded === 0 ? '0'.repeat(places) : '';
  const integerText = placeInteger(integerDigits, integerTokens, thousands, commaRole);
  let text = integerText;
  if (pointAt !== -1) {
    text += '.' + placeDecimals(fraction, decimalTokens);
  }
  return sign + text + exponentText;
}

function lastIndexOf<T>(items: readonly T[], test: (item: T) => boolean): number {
  for (let at = items.length - 1; at >= 0; at--) {
    if (test(items[at])) {
      return at;
    }
  }
  return -1;
}

/** Rounded to a number of places, a tie away from zero, at fifteen significant digits. */
function roundTo(value: number, places: number): number {
  const factor = 10 ** places;
  const scaled = Number((value * factor).toPrecision(15));
  return Math.round(scaled) / factor;
}

/**
 * The integer part, laid into its placeholders from the right.
 *
 * Digits fill the placeholders right to left, and any the placeholders
 * cannot hold go to the leftmost; a placeholder with no digit left is
 * a `0`, nothing for `#`, a space for `?`. Literals between them stay
 * where they were written. With a thousands separator, the separators
 * are placed by the digits' positions, and one that falls among the
 * padding is padding itself.
 */
function placeInteger(
  digits: string,
  tokens: readonly Token[],
  thousands: boolean,
  commaRole: ReadonlyMap<number, 'thousands' | 'scale' | 'literal'>
): string {
  const slots = tokens.map((token, at) => ({ token, at })).filter(({ token }) => token.kind === 'digit');
  const count = slots.length;
  const padding: string[] = slots.map(({ token }) => (token as Extract<Token, { kind: 'digit' }>).char);
  // One character per slot, right-aligned; the first slot takes the overflow.
  const cells: string[] = new Array<string>(count).fill('');
  for (let at = 0; at < count; at++) {
    const fromRight = count - 1 - at;
    const digitAt = digits.length - 1 - fromRight;
    if (digitAt >= 0) {
      cells[at] = at === 0 ? digits.slice(0, digitAt + 1) : digits[digitAt];
    } else {
      cells[at] = padding[at] === '0' ? '0' : padding[at] === '?' ? ' ' : '';
    }
  }
  if (count === 0) {
    return tokens.map(token => (token.kind === 'literal' ? token.text : '')).join('');
  }
  if (thousands) {
    // Grouped over the whole run: a digit string the width of the
    // placeholders or of the number, whichever is wider.
    const flat = cells.join('');
    const grouped: string[] = [];
    const chars = [...flat];
    for (let at = 0; at < chars.length; at++) {
      const fromRight = chars.length - at;
      grouped.push(chars[at]);
      if (fromRight > 1 && (fromRight - 1) % 3 === 0) {
        const left = chars[at];
        grouped.push(left === ' ' ? ' ' : left === '' ? '' : ',');
      }
    }
    const before = tokens.slice(0, slots[0].at).map(token => (token.kind === 'literal' ? token.text : '')).join('');
    const after = tokens
      .slice(slots[count - 1].at + 1)
      .map((token, at) => (token.kind === 'literal' ? token.text : token.kind === 'comma' && commaRole.get(slots[count - 1].at + 1 + at) === 'literal' ? ',' : ''))
      .join('');
    return before + grouped.join('') + after;
  }
  let text = '';
  let slot = 0;
  tokens.forEach((token, at) => {
    if (token.kind === 'digit') {
      text += cells[slot++];
    } else if (token.kind === 'literal') {
      text += token.text;
    } else if (token.kind === 'comma' && commaRole.get(at) === 'literal') {
      text += ',';
    } else if (token.kind === 'percent') {
      text += '%';
    }
  });
  return text;
}

/** The decimals, left to right; trailing zeros under `#` are dropped and under `?` are spaces. */
function placeDecimals(digits: string, tokens: readonly Token[]): string {
  const places = tokens.filter(token => token.kind === 'digit') as Extract<Token, { kind: 'digit' }>[];
  let significant = digits.length;
  while (significant > 0 && digits[significant - 1] === '0' && places[significant - 1]?.char !== '0') {
    significant--;
  }
  let text = '';
  let slot = 0;
  for (const token of tokens) {
    if (token.kind === 'digit') {
      const at = slot++;
      text += at < significant ? digits[at] : token.char === '0' ? '0' : token.char === '?' ? ' ' : '';
    } else if (token.kind === 'literal') {
      text += token.text;
    } else if (token.kind === 'percent') {
      text += '%';
    } else if (token.kind === 'comma') {
      // A scaling comma among the decimals prints nothing.
    }
  }
  return text;
}

/**
 * A fraction: `# ?/?`, `# ??/??`, `?/8`.
 *
 * The numerator is the run of placeholders just before the slash, and
 * anything earlier is the whole-number part, laid out like any integer.
 * With a whole part, it takes the whole number and the fraction what is
 * left; without one, the fraction is everything. The denominator is
 * written (`/8`) or the closest one its placeholders can hold.
 *
 * A zero whole under `#` drops what sits between it and the numerator,
 * under `?` makes it spaces, and under `0` prints it. A zero fraction,
 * unless the numerator is `0`, drops everything from there to the
 * denominator's end — or, with a `?` anywhere, is spaces the same
 * width, so a column of them still lines up.
 */
function formatFraction(value: number, tokens: readonly Token[], slash: number): string {
  const before = tokens.slice(0, slash);
  const after = tokens.slice(slash + 1);
  let numeratorStart = before.length - 1;
  while (numeratorStart >= 0 && before[numeratorStart].kind !== 'digit') {
    numeratorStart--;
  }
  const numeratorEnd = numeratorStart + 1;
  while (numeratorStart > 0 && before[numeratorStart - 1].kind === 'digit') {
    numeratorStart--;
  }
  const lastWhole = lastIndexOf(before.slice(0, numeratorStart), token => token.kind === 'digit');
  const hasWhole = lastWhole !== -1;
  const wholeTokens = before.slice(0, lastWhole + 1);
  // Literals ahead of a code with no whole part are simply printed.
  const lead = hasWhole ? '' : literalsOf(before.slice(0, numeratorStart));
  const between = hasWhole ? before.slice(lastWhole + 1, numeratorStart) : [];
  const numeratorTokens = before.slice(numeratorStart, numeratorEnd);
  const nearSlash = literalsOf(before.slice(numeratorEnd));
  const denominatorFirst = after.findIndex(token => token.kind === 'digit' || (token.kind === 'literal' && /^\d$/.test(token.text)));
  const denominatorLast = lastIndexOf(after, token => token.kind === 'digit' || (token.kind === 'literal' && /^\d$/.test(token.text)));
  const afterSlash = literalsOf(after.slice(0, denominatorFirst));
  const denominatorTokens = after.slice(denominatorFirst, denominatorLast + 1);
  const trailing = literalsOf(after.slice(denominatorLast + 1));
  const fixed = denominatorTokens.every(token => token.kind === 'literal')
    ? denominatorTokens.map(token => (token as { text: string }).text).join('')
    : '';
  const denominatorPlaces = denominatorTokens.filter(token => token.kind === 'digit') as Extract<Token, { kind: 'digit' }>[];

  let whole = hasWhole ? Math.floor(value) : 0;
  const part = hasWhole ? value - whole : value;
  let numerator: number;
  let denominator: number;
  if (fixed !== '') {
    denominator = Number(fixed);
    numerator = Math.round(part * denominator);
  } else {
    [numerator, denominator] = bestFraction(part, 10 ** Math.max(1, denominatorPlaces.length) - 1);
  }
  if (hasWhole && numerator === denominator) {
    whole++;
    numerator = 0;
  }

  const charOf = (items: readonly Token[]): string => {
    const last = lastIndexOf(items, token => token.kind === 'digit');
    return last === -1 ? '' : (items[last] as Extract<Token, { kind: 'digit' }>).char;
  };
  const width = (items: readonly Token[]): number =>
    items.reduce((sum, token) => sum + (token.kind === 'literal' ? token.text.length : token.kind === 'digit' ? 1 : 0), 0);

  if (hasWhole && numerator === 0 && charOf(numeratorTokens) !== '0') {
    const wholeText = placeInteger(String(whole), wholeTokens, false, new Map());
    const padded = [...wholeTokens, ...numeratorTokens, ...denominatorTokens].some(token => token.kind === 'digit' && token.char === '?');
    const gap = width(between) + width(before.slice(numeratorStart)) + 1 + width(after.slice(0, denominatorLast + 1));
    return wholeText + (padded ? ' '.repeat(gap) : '') + trailing;
  }
  let wholeText = '';
  if (hasWhole) {
    const blank = charOf(wholeTokens);
    wholeText = placeInteger(whole === 0 ? '' : String(whole), wholeTokens, false, new Map());
    wholeText += whole !== 0 || blank === '0' ? literalsOf(between) : blank === '?' ? ' '.repeat(width(between)) : '';
  }
  const numeratorText = placeInteger(String(numerator), numeratorTokens, false, new Map());
  let denominatorText = fixed;
  if (fixed === '') {
    const zeros = denominatorPlaces.filter(token => token.char === '0').length;
    const spaces = denominatorPlaces.filter(token => token.char !== '#').length;
    denominatorText = String(denominator).padStart(zeros, '0').padEnd(spaces, ' ');
  }
  return `${lead}${wholeText}${numeratorText}${nearSlash}/${afterSlash}${denominatorText}${trailing}`;
}

function literalsOf(tokens: readonly Token[]): string {
  return tokens.map(token => (token.kind === 'literal' ? token.text : token.kind === 'percent' ? '%' : '')).join('');
}

/** The closest fraction with a denominator no larger than `limit`. */
function bestFraction(value: number, limit: number): [number, number] {
  let best: [number, number] = [Math.round(value), 1];
  let error = Math.abs(value - best[0]);
  for (let denominator = 1; denominator <= limit; denominator++) {
    const numerator = Math.round(value * denominator);
    const off = Math.abs(value - numerator / denominator);
    if (off < error - 1e-12) {
      best = [numerator, denominator];
      error = off;
    }
  }
  return best;
}

// ---------------------------------------------------------------------------
// Dates and times
// ---------------------------------------------------------------------------

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

/** Where 1904's serial 0 falls in the 1900 system. */
const SHIFT_1904 = 1462;

function formatDate(value: number, tokens: readonly Token[], options: FormatOptions): string | CellError {
  const elapsed = tokens.some(token => token.kind === 'elapsed');
  if (value < 0 && !elapsed) {
    return VALUE;
  }
  const sign = value < 0 ? '-' : '';
  const magnitude = Math.abs(value);

  // Rounded to the smallest unit shown: thousandths of a second, or
  // hundredths, tenths, or whole seconds.
  const fractionPlaces = fractionalSeconds(tokens);
  const unit = 10 ** fractionPlaces;
  const totalUnits = Math.round(Number((magnitude * 86_400 * unit).toPrecision(15)));
  const days = Math.floor(totalUnits / (86_400 * unit));
  const inDay = totalUnits - days * 86_400 * unit;
  const seconds = Math.floor(inDay / unit);
  const fraction = inDay - seconds * unit;
  const hours = Math.floor(seconds / 3600);
  const minutes = Math.floor((seconds % 3600) / 60);
  const secs = seconds % 60;

  const serial = options.date1904 === true ? days + SHIFT_1904 : days;
  const date = dateOfSerial(serial);
  const weekday = weekdayOf(serial);
  const twelve = tokens.some(token => token.kind === 'ampm');

  // `m` is minutes beside an hour or a second, and a month otherwise.
  const units = tokens
    .map((token, at) => ({ at, unit: token.kind === 'date' ? token.code[0] : token.kind === 'elapsed' ? token.unit : '' }))
    .filter(entry => entry.unit !== '');
  const isMinute = (index: number): boolean => {
    const position = units.findIndex(entry => entry.at === index);
    const previous = units[position - 1]?.unit;
    const next = units[position + 1]?.unit;
    return previous === 'h' || previous === 's' || next === 's';
  };
  // Once a section counts elapsed hours, a plain `h` in it counts them
  // too; minutes and seconds likewise.
  const elapsedUnits = new Set(tokens.flatMap(token => (token.kind === 'elapsed' ? [token.unit] : [])));
  const totalHours = days * 24 + hours;
  const totals = { h: totalHours, m: totalHours * 60 + minutes, s: (totalHours * 60 + minutes) * 60 + secs };

  let text = '';
  for (let at = 0; at < tokens.length; at++) {
    const token = tokens[at];
    switch (token.kind) {
      case 'literal':
        text += token.text;
        break;
      case 'digit':
        // `0` after seconds and a point is a fraction of a second, which
        // `fractionalSeconds` has already counted.
        break;
      case 'point': {
        const places = countAfter(tokens, at);
        if (places > 0) {
          text += '.' + String(fraction).padStart(fractionPlaces, '0').slice(0, places);
        } else {
          text += '.';
        }
        break;
      }
      case 'elapsed': {
        text += String(totals[token.unit]).padStart(token.width, '0');
        break;
      }
      case 'ampm': {
        const afternoon = hours >= 12;
        if (token.code.length === 5) {
          text += afternoon ? 'PM' : 'AM';
        } else {
          const letter = afternoon ? token.code[2] : token.code[0];
          text += letter;
        }
        break;
      }
      case 'date': {
        const code = token.code;
        const letter = code[0];
        if (letter === 'y') {
          text += code.length <= 2 ? String(date.year % 100).padStart(2, '0') : String(date.year);
        } else if (letter === 'd') {
          text +=
            code.length === 1
              ? String(date.day)
              : code.length === 2
                ? String(date.day).padStart(2, '0')
                : code.length === 3
                  ? DAYS[weekday - 1].slice(0, 3)
                  : DAYS[weekday - 1];
        } else if ((letter === 'h' || letter === 's') && elapsedUnits.has(letter)) {
          text += String(totals[letter]).padStart(code.length, '0');
        } else if (letter === 'm' && elapsedUnits.has('m') && isMinute(at) && code.length <= 2) {
          text += String(totals.m).padStart(code.length, '0');
        } else if (letter === 'h') {
          const shown = twelve ? (hours % 12 === 0 ? 12 : hours % 12) : hours;
          text += code.length === 1 ? String(shown) : String(shown).padStart(2, '0');
        } else if (letter === 's') {
          text += code.length === 1 ? String(secs) : String(secs).padStart(2, '0');
        } else if (letter === 'm') {
          if (isMinute(at) && code.length <= 2) {
            text += code.length === 1 ? String(minutes) : String(minutes).padStart(2, '0');
          } else {
            text +=
              code.length === 1
                ? String(date.month)
                : code.length === 2
                  ? String(date.month).padStart(2, '0')
                  : code.length === 3
                    ? MONTHS[date.month - 1].slice(0, 3)
                    : code.length === 5
                      ? MONTHS[date.month - 1][0]
                      : MONTHS[date.month - 1];
          }
        }
        break;
      }
      default:
        break;
    }
  }
  return sign + text;
}

/** How many zeros follow the point after the seconds, which is how finely time is shown. */
function fractionalSeconds(tokens: readonly Token[]): number {
  for (let at = 0; at < tokens.length; at++) {
    if (tokens[at].kind === 'point') {
      const places = countAfter(tokens, at);
      if (places > 0) {
        return Math.min(places, 3);
      }
    }
  }
  return 0;
}

function countAfter(tokens: readonly Token[], at: number): number {
  let count = 0;
  for (let next = at + 1; next < tokens.length && tokens[next].kind === 'digit'; next++) {
    count++;
  }
  return count;
}
