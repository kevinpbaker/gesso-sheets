import { describe, expect, it } from 'vitest';

import { darkColors, lightColors } from 'gesso-core';

import { inkFor } from './cellInk';

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
