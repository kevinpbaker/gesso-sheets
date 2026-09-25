import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

import { describe, expect, it } from 'vitest';

import { ICONS } from './icons';
import * as heroicons from './heroicons';

/**
 * That the glyphs in the bundle are still the glyphs in the package.
 *
 * `src/app/heroicons.ts` is generated, and generated files rot in a
 * particular way: `pnpm up heroicons` moves the package and nothing
 * says the checked-in copy is now a version behind. Nobody would
 * notice — an icon one release out of date still draws — until the
 * day a glyph is redrawn upstream and the toolbar is the only place
 * in the world still showing the old one.
 *
 * So the file is checked against the package rather than trusted: each
 * glyph is lifted again here, by reading the SVG directly, and has to
 * match. This deliberately does not call `scripts/icons.ts` — a
 * generator compared against itself agrees with itself. Failing means
 * running `pnpm icons` and committing what it writes.
 */
const HERE = dirname(fileURLToPath(import.meta.url));
const PACKAGE = join(HERE, '..', '..', 'node_modules', 'heroicons');

/** The generated file says which glyph each constant came from. */
const LIFTED = /\/\*\* Heroicons `([a-z0-9-]+)`, 24×24 (outline|solid)\. \*\/\nexport const (\w+):/g;

interface Source {
  readonly name: string;
  readonly set: string;
  readonly constant: string;
}

function generated(): string {
  return readFileSync(join(HERE, 'heroicons.ts'), 'utf8');
}

function sources(): Source[] {
  return [...generated().matchAll(LIFTED)].map(match => ({
    name: match[1]!,
    set: match[2]!,
    constant: match[3]!
  }));
}

/** The `d` of every path in the file, in order, as `Icon` wants it. */
function pathOf(source: Source): string {
  const svg = readFileSync(join(PACKAGE, '24', source.set, `${source.name}.svg`), 'utf8');
  return [...svg.matchAll(/\sd="([^"]+)"/g)].map(match => match[1]!).join('');
}

describe('the icon set', () => {
  it('has glyphs to check', () => {
    // A guard against the checks below passing because they found nothing.
    expect(sources().length).toBeGreaterThan(8);
  });

  it('draws what the installed heroicons draws', () => {
    const stale: string[] = [];
    for (const source of sources()) {
      const glyph = (heroicons as Record<string, unknown>)[source.constant] as { path: string };
      if (glyph.path !== pathOf(source)) {
        stale.push(source.name);
      }
    }
    expect(stale, 'run `pnpm icons` to lift these again').toEqual([]);
  });

  it('was generated from the installed version', () => {
    const version = JSON.parse(readFileSync(join(PACKAGE, 'package.json'), 'utf8')).version as string;
    expect(generated()).toContain(`heroicons@${version}`);
  });

  it('paints every glyph the way the set is authored', () => {
    for (const [name, glyph] of Object.entries(ICONS)) {
      expect(glyph.path.length, name).toBeGreaterThan(0);
      // The rasteriser is handed viewBox units, so a glyph on another
      // grid would draw at the wrong size rather than fail.
      expect(glyph.viewBox, name).toBe(24);
      expect(glyph.style, name).toBe('stroke');
      expect(glyph.strokeWidth, name).toBe(1.5);
    }
  });

  /**
   * The one drawn path, kept honest.
   *
   * `alignCenter` is this application's own — Heroicons has no centred
   * glyph — and the risk is that a second one, and a third, arrive
   * beside it without the same conversation. Everything else has to
   * come from the package, which is what makes the set *the* set.
   */
  it('uses the package for everything but the centred bars', () => {
    const lifted = new Set(sources().map(source => source.constant));
    const own = Object.entries(ICONS).filter(
      ([, glyph]) => !lifted.has(constantFor(glyph.path))
    );
    expect(own.map(([name]) => name)).toEqual(['alignCenter']);
  });
});

/** Which generated constant holds this path, or `''` when none does. */
function constantFor(path: string): string {
  for (const [constant, glyph] of Object.entries(heroicons)) {
    if ((glyph as { path: string }).path === path) {
      return constant;
    }
  }
  return '';
}
