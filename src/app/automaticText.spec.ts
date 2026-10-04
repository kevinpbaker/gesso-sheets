import { describe, expect, it } from 'vitest';

import { automaticText } from './automaticText';

describe('automaticText', () => {
  it('is the theme text with no fill', () => {
    expect(automaticText('')).toBe('text');
  });

  it('is dark on a pale fill and light on a deep one, in either theme', () => {
    expect(automaticText('#fce8e6')).toBe('#000000');
    expect(automaticText('#fff')).toBe('#000000');
    expect(automaticText('#1f3864')).toBe('#ffffff');
    expect(automaticText('#000000')).toBe('#ffffff');
  });

  it('keeps the theme text for a fill it cannot measure', () => {
    expect(automaticText('red')).toBe('text');
  });
});
