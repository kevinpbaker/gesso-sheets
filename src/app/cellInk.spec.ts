import { describe, expect, it } from 'vitest';

import { darkColors, lightColors } from 'gesso-core';

import { inkFor, shownFill, shownPalette } from './cellInk';
import { SCENARIO_CHANGED, SCENARIO_TYPED } from './scenarioPaint';
import { PLAIN_PAINT } from './SheetContract';

describe('inkFor', () => {
  it('is the theme ink on a cell with no fill, so the sheet follows light and dark', () => {
    expect(inkFor('')).toBe('text');
  });

  it('is the light palette ink on a pale fill, in either theme', () => {
    // The seed's header fill and a colour scale's palest step.
    expect(inkFor('#f1f3f4')).toBe(lightColors.text);
    expect(inkFor('#fce8e6')).toBe(lightColors.text);
  });

  it('is the dark palette ink on a dark fill', () => {
    expect(inkFor('#1f3864')).toBe(darkColors.text);
    expect(inkFor('#000000')).toBe(darkColors.text);
  });

  it('falls back to the theme ink for a fill it cannot read or can see through', () => {
    expect(inkFor('not a colour')).toBe('text');
    expect(inkFor('rgba(0, 0, 0, 0.1)')).toBe('text');
  });
});

/**
 * A light fill on a dark sheet, drawn dark.
 *
 * The seed's pale header rows and its colour scale's pastels were
 * bright stripes across the dark theme, which drew the author's
 * colour exactly as written.
 */
describe('shownFill', () => {
  it('is the author\'s colour in the light theme, always', () => {
    expect(shownFill('#f1f3f4', false)).toBe('#f1f3f4');
    expect(shownFill('#fce8e6', false)).toBe('#fce8e6');
  });

  it('darkens a light fill in the dark theme, keeping its hue', () => {
    // A pale grey header, a pale red step of a colour scale, the seed's input yellow.
    expect(shownFill('#f1f3f4', true)).toBe('#272c2e');
    expect(shownFill('#fce8e6', true)).toBe('#401b17');
    expect(shownFill('#fff8d6', true)).toBe('#4e4413');
  });

  it('leaves a fill that is not light as it was written', () => {
    // The navy title bar, a full blue, a highlighter yellow.
    expect(shownFill('#1f3864', true)).toBe('#1f3864');
    expect(shownFill('#4285f4', true)).toBe('#4285f4');
    expect(shownFill('#ffd966', true)).toBe('#ffd966');
    expect(shownFill('', true)).toBe('');
  });

  it('keeps the scenario\'s two tints two, and apart from the input fill', () => {
    const typed = shownFill(SCENARIO_TYPED, true);
    const changed = shownFill(SCENARIO_CHANGED, true);
    expect([typed, changed]).toEqual(['#291b61', '#1c1341']);
    expect(new Set([typed, changed, shownFill('#fff8d6', true)]).size).toBe(3);
  });

  it('is drawn over in the ink that reads on what is shown', () => {
    expect(inkFor(shownFill('#f1f3f4', true))).toBe(darkColors.text);
    expect(inkFor(shownFill(SCENARIO_TYPED, true))).toBe(darkColors.text);
    expect(inkFor(shownFill('#1f3864', true))).toBe(darkColors.text);
  });
});

describe('shownPalette', () => {
  const paint = (fill: string, color = '') => ({ ...PLAIN_PAINT, fill, color });

  it('is the palette it was given in the light theme', () => {
    const entries = [PLAIN_PAINT, paint('#fce8e6', '#c5221f')];
    expect(shownPalette(entries, false)).toBe(entries);
  });

  it('keeps an entry it does not change, so nothing is pushed for it', () => {
    const navy = paint('#1f3864', '#ffffff');
    expect(shownPalette([PLAIN_PAINT, navy], true)[1]).toBe(navy);
  });

  /** A highlight is a pair, and dark red on dark red is no highlight. */
  it('turns an author\'s dark text light where it darkened the fill', () => {
    expect(shownPalette([paint('#fce8e6', '#c5221f'), paint('#e6f4ea', '#137333')], true).map(entry => [entry.fill, entry.color])).toEqual([
      ['#401b17', '#ef9998'],
      ['#243a2a', '#b1f2c7']
    ]);
  });
});
