/**
 * Tessera's own icon: a mosaic of nine tiles, one of them lit.
 *
 *   pnpm app-icon
 *
 * A tessera is one tile of a mosaic, and a sheet is a mosaic of cells,
 * so the icon is three rows of three tiles set in dark grout, laid a
 * little off square the way a hand lays them, with the middle one gold
 * for the selected cell. At 16 px the tilt is gone and what is left is
 * a grid with one cell picked out, which is what a spreadsheet's icon
 * has to say at that size.
 *
 * This writes `public/favicon.svg` and renders the PNGs the manifest
 * and `apple-touch-icon` name with headless Chrome, so changing a
 * colour here and running it again is the whole of changing the icon.
 * Set `CHROME_BIN` to choose the binary.
 */
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { findChrome } from './lib/devtools.ts';

const PUBLIC = join(dirname(fileURLToPath(import.meta.url)), '..', 'public');

const GROUT = '#1e2a44';
const LIT = '#f5b942';
/** Row by row; `null` is the lit tile. */
const TILES: (string | null)[] = ['#5b82c9', '#4a6fb5', '#6f93d4', '#3f63a6', null, '#5579bf', '#6a8dcf', '#4568ad', '#5e85cb'];
/** Degrees each tile is turned by, so the grid reads as laid, not drawn. */
const TILT = [-2.5, 1.5, -1, 2, 0, -2, 1, -1.5, 2.5];

interface Shape {
  /** Corner radius of the background; 0 for a platform that rounds it itself. */
  corner: number;
  /** Side of the tile grid in a 64-unit square. */
  grid: number;
}

/** The ordinary icon, the full-bleed one iOS rounds, and the maskable one whose tiles stay inside the safe circle. */
const ICON: Shape = { corner: 14, grid: 44 };
const SQUARE: Shape = { corner: 0, grid: 44 };
const MASKABLE: Shape = { corner: 0, grid: 34 };

function svg({ corner, grid }: Shape, size?: number): string {
  const gap = grid * 0.075;
  const side = (grid - 2 * gap) / 3;
  const start = (64 - grid) / 2;
  const tiles = TILES.map((fill, i) => {
    const x = start + (i % 3) * (side + gap);
    const y = start + Math.floor(i / 3) * (side + gap);
    const centre = `${(x + side / 2).toFixed(2)} ${(y + side / 2).toFixed(2)}`;
    return `  <rect x="${x.toFixed(2)}" y="${y.toFixed(2)}" width="${side.toFixed(2)}" height="${side.toFixed(2)}" rx="1.6" fill="${fill ?? LIT}" transform="rotate(${TILT[i]} ${centre})"/>`;
  });
  const dimensions = size === undefined ? '' : ` width="${size}" height="${size}"`;
  return [
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64"${dimensions}>`,
    `  <rect width="64" height="64" rx="${corner}" fill="${GROUT}"/>`,
    ...tiles,
    '</svg>',
    '',
  ].join('\n');
}

function render(shape: Shape, size: number, name: string, chrome: string, scratch: string): void {
  const source = join(scratch, `${name}.svg`);
  const target = join(PUBLIC, name);
  writeFileSync(source, svg(shape, size));
  execFileSync(chrome, [
    '--headless',
    '--disable-gpu',
    '--hide-scrollbars',
    '--default-background-color=00000000',
    `--window-size=${size},${size}`,
    `--screenshot=${target}`,
    `file://${source}`,
  ], { stdio: 'ignore' });
  if (readFileSync(target).length === 0) {
    throw new Error(`Chrome wrote nothing to ${target}`);
  }
  console.log(`wrote public/${name}`);
}

writeFileSync(join(PUBLIC, 'favicon.svg'), svg(ICON));
console.log('wrote public/favicon.svg');

const chrome = findChrome();
const scratch = mkdtempSync(join(tmpdir(), 'tessera-icon-'));
try {
  render(ICON, 192, 'icon-192.png', chrome, scratch);
  render(ICON, 512, 'icon-512.png', chrome, scratch);
  render(MASKABLE, 512, 'icon-maskable-512.png', chrome, scratch);
  render(SQUARE, 180, 'apple-touch-icon.png', chrome, scratch);
} finally {
  rmSync(scratch, { recursive: true, force: true });
}
