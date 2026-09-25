/**
 * Heroicons, turned into something a canvas can draw.
 *
 *   pnpm icons
 *
 * The `heroicons` package ships SVG files and React and Vue
 * components, and this application can use none of the three: there is
 * no DOM under the interface, so there is nothing to hand an `<svg>`
 * to. What Gesso's `Icon` wants is the one part of the file that is
 * not a document — the `d` of each path, the square it was drawn in,
 * and whether it is filled or stroked. This script lifts exactly that
 * into `src/app/heroicons.ts`, so the dependency is a development one
 * and the bundle carries a few hundred bytes of path data rather than
 * a component library.
 *
 * Generating rather than hand-copying is for the upgrade: `pnpm up
 * heroicons && pnpm icons` re-lifts every glyph, and a path that
 * changed shape upstream shows up as a diff instead of as a drawing
 * that quietly stopped matching the set.
 *
 * `WANTED` below is the whole of what this application uses. Adding an
 * icon means adding a line here and running it again; nothing scans
 * the source for uses, because a list that can be read in one screen
 * is worth more here than one that cannot be wrong.
 */
import { readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

/** Heroicon names, in the 24×24 set, as the package files are called. */
const WANTED = [
  'arrow-path',
  'arrow-uturn-left',
  'arrow-uturn-right',
  'bars-3-bottom-left',
  'bars-3-bottom-right',
  'bold',
  'currency-dollar',
  'italic',
  'percent-badge',
  'underline'
] as const;

/**
 * Which set to lift from.
 *
 * Outline, for all of it. A toolbar of solid glyphs at sixteen pixels
 * reads as a row of blobs, and the stroked set is what the rasteriser
 * is already set up for — it draws with round caps and joins, which is
 * how Heroicons' outline set is authored.
 */
const SET = 'outline';

const here = dirname(fileURLToPath(import.meta.url));
const root = join(here, '..');
const from = join(root, 'node_modules', 'heroicons', '24', SET);
const to = join(root, 'src', 'app', 'heroicons.ts');

function main(): void {
  const version = JSON.parse(
    readFileSync(join(root, 'node_modules', 'heroicons', 'package.json'), 'utf8')
  ).version as string;

  const glyphs = WANTED.map(name => lift(name));
  writeFileSync(to, file(version, glyphs), 'utf8');
  console.log(`  ${glyphs.length} glyphs from heroicons@${version} → src/app/heroicons.ts`);
}

interface Lifted {
  readonly name: string;
  readonly constant: string;
  readonly path: string;
  readonly viewBox: number;
  readonly style: 'fill' | 'stroke';
  readonly strokeWidth: number;
  readonly fillRule: 'nonzero' | 'evenodd' | undefined;
}

/**
 * One SVG file, as the six things an icon is.
 *
 * The style comes from the file rather than from the directory it is
 * in, because the file says so and a directory is a convention: an
 * outline icon strokes `currentColor` over `fill="none"`, a solid one
 * fills it.
 *
 * **Several paths become one.** `IconProps.path` is a single string,
 * and a glyph drawn as three `<path>` elements is three subpaths of
 * one shape — concatenating their `d` is what an SVG renderer is
 * doing anyway. It holds because Heroicons gives every path in a file
 * the same paint: the same stroke width for the outline set, the same
 * `fill-rule` for the solid one. A set that varied paint per path
 * could not be lifted this way, and this would have to say so.
 */
function lift(name: string): Lifted {
  const svg = readFileSync(join(from, `${name}.svg`), 'utf8');
  const paths = [...svg.matchAll(/\sd="([^"]+)"/g)].map(match => match[1]!);
  if (paths.length === 0) {
    throw new Error(`${name}.svg has no path data.`);
  }

  const viewBox = attribute(svg, 'viewBox');
  if (viewBox === undefined || !/^0 0 (\d+) \1$/.test(viewBox)) {
    throw new Error(`${name}.svg is not drawn in a square starting at the origin: ${viewBox}.`);
  }

  const stroked = attribute(svg, 'stroke') === 'currentColor';
  return {
    name,
    constant: camel(name),
    path: paths.join(''),
    viewBox: Number(viewBox.split(' ')[2]),
    style: stroked ? 'stroke' : 'fill',
    strokeWidth: stroked ? Number(attribute(svg, 'stroke-width') ?? '1.5') : 0,
    /**
     * `nonzero` is the canvas default and `evenodd` is what the solid
     * set declares; a glyph with a hole in it drawn under the wrong
     * one is a filled blob, which is the failure `IconSpec.fillRule`
     * exists to prevent.
     */
    fillRule: stroked ? undefined : attribute(svg, 'fill-rule') === 'evenodd' ? 'evenodd' : 'nonzero'
  };
}

function attribute(svg: string, name: string): string | undefined {
  return new RegExp(`\\s${name}="([^"]*)"`).exec(svg)?.[1];
}

function camel(name: string): string {
  return name.replace(/-(.)/g, (_, character: string) => character.toUpperCase());
}

function file(version: string, glyphs: readonly Lifted[]): string {
  const declarations = glyphs
    .map(glyph => {
      const path = `'${glyph.path}'`;
      const fields = [
        /** Long path data on its own line, so the fields stay readable beside it. */
        path.length > 72 ? `path:\n    ${path},` : `path: ${path},`,
        `viewBox: ${glyph.viewBox},`,
        `style: '${glyph.style}',`,
        ...(glyph.style === 'stroke' ? [`strokeWidth: ${glyph.strokeWidth}`] : [`fillRule: '${glyph.fillRule}'`])
      ];
      return (
        `/** Heroicons \`${glyph.name}\`, 24×24 ${SET}. */\n` +
        `export const ${glyph.constant}: Glyph = {\n  ${fields.join('\n  ')}\n};`
      );
    })
    .join('\n\n');

  return `/**
 * Heroicons, as path data.
 *
 * Generated by \`pnpm icons\` from heroicons@${version}. Do not edit:
 * \`scripts/icons.ts\` holds the list of glyphs and rewrites this file.
 *
 * Heroicons is MIT licensed — Copyright (c) Tailwind Labs, Inc. The
 * licence is at \`node_modules/heroicons/LICENSE\` and travels with
 * this data.
 */
import type { Glyph } from './icons';

${declarations}
`;
}

main();
