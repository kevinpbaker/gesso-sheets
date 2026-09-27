import { describe, expect, it } from 'vitest';

import { parseAddress } from '../sheet/A1';
import { Workbook } from '../sheet/Workbook';
import { ConditionalPainter } from './ConditionalPaint';

/**
 * A colour scale over cells an array spilled into.
 *
 * The scale's extent was found by walking the cells that hold
 * something, and a spilled cell holds nothing of its own — so over a
 * column that one formula filled, the scale saw one value and painted
 * every cell the middle colour. Found by looking at the seeded
 * forecast, twelve rising months all the same pale yellow.
 */
describe('a colour scale', () => {
  it('spreads over the cells a formula spilled into, as over typed ones', () => {
    const sheet = new Workbook().sheet(0);
    sheet.setCell(0, 0, '=SEQUENCE(5)');
    sheet.recalculate();
    const painter = new ConditionalPainter(() => sheet);
    painter.setRules([{ range: parseAddress('A1:A5')!, scale: { from: '#ff0000', middle: '#ffff00', to: '#00ff00' }, test: null }]);
    const fill = (row: number) => painter.paintFor(row, 0, sheet.value(row, 0))?.fill;
    expect(fill(0)).toBe('#ff0000');
    expect(fill(4)).toBe('#00ff00');
    expect(fill(2)).not.toBe(fill(0));
  });
});
