import { afterEach, describe, expect, it } from 'vitest';

import { renderTest, serveForTest, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';
import { createComponent } from 'gesso-framework';
import type { PaintBox, PaintSurface, UiKeyModifiers } from 'gesso-core';

import { relativeRef } from '../sheet/A1';
import { SERIES_COLOURS } from '../sheet/Chart';
import { DARK_INK, LIGHT_INK, readableOnDark, seriesColour } from './ChartPaint';
import { SheetApp } from './SheetApp';
import { SheetDocument } from './SheetDocument';
import { sheetChannel } from './sheetChannel';
import { SheetService } from './SheetService';

/**
 * A chart in the dark theme is a dark panel.
 *
 * A chart is a `Paint`: a picture made once and kept until one of its
 * inputs changes. The theme was not one of them, so the picture made in
 * the light theme — a pale panel, dark ink, pale gridlines — went on
 * being shown on the dark sheet after View ▸ Theme went dark.
 */

/** Every colour a picture asks for, in order, and nothing else of it. */
function colours(draw: (surface: PaintSurface, box: PaintBox) => void): { fills: unknown[]; strokes: unknown[] } {
  const fills: unknown[] = [];
  const strokes: unknown[] = [];
  const surface = new Proxy(
    {},
    {
      get: (_target, name) => (value: unknown) => {
        if (name === 'fillColor') {
          fills.push(value);
        } else if (name === 'strokeColor') {
          strokes.push(value);
        }
      }
    }
  ) as PaintSurface;
  draw(surface, { width: 320, height: 200, paddingTop: 0, paddingRight: 0, paddingBottom: 0, paddingLeft: 0, scale: 1 } as PaintBox);
  return { fills, strokes };
}

describe('a chart in either theme', () => {
  let h: { ui: Rendered; served: ServedForTest };

  afterEach(() => {
    h?.ui.unmount();
    h?.served.dispose();
  });

  async function settle(): Promise<void> {
    await h.ui.settle();
    await h.served.settle();
    await h.ui.settle();
  }

  async function press(key: string, modifiers: Partial<UiKeyModifiers> = {}): Promise<void> {
    h.ui.fireEvent.press(key, modifiers);
    await settle();
  }

  async function mount(): Promise<void> {
    const document = new SheetDocument();
    const rows = [
      ['Month', 'Units', 'Cost'],
      ['Jul', '5', '50'],
      ['Aug', '10', '90'],
      ['Sep', '15', '120']
    ];
    rows.forEach((line, row) => line.forEach((value, column) => document.setCell(row, column, value)));
    document.addChart({
      kind: 'column',
      title: 'Units',
      legend: true,
      range: { start: relativeRef(0, 0), end: relativeRef(3, 2) },
      place: { x: 420, y: 20, width: 320, height: 200 }
    });
    document.sheet.recalculate();
    const service = new SheetService(document, { rowCount: 100, columnCount: 10 });
    await service.restore();
    const served = serveForTest([sheetChannel(service)]);
    const ui = renderTest(createComponent(SheetApp), { channels: served.registry, width: 1100, height: 700 });
    h = { ui, served };
    await settle();
  }

  /** The chart's picture as the engine holds it: what it draws, and what makes it draw again. */
  const picture = () =>
    h.ui.getByRole('image', { name: /Units, chart of/ }).properties.get('paint') as {
      draw: (surface: PaintSurface, box: PaintBox) => void;
      inputs: readonly unknown[];
    };

  it('is drawn in the theme\'s own colours in the light theme, as it always was', async () => {
    await mount();
    const { fills, strokes } = colours(picture().draw);
    expect(fills.slice(0, 2)).toEqual(['surface', 'text']);
    // The panel's edge, the zero line, and three gridlines above it.
    expect(strokes).toEqual(['border', 'border', 'controlBorder', 'controlBorder', 'controlBorder']);
    expect(fills).toContain(SERIES_COLOURS[0]);
    expect(fills).toContain('textMuted');
  });

  it('is made again for the dark theme the moment View ▸ Theme goes dark', async () => {
    await mount();
    expect(picture().inputs.at(-1)).toBe(false);

    h.ui.fireEvent.focus(h.ui.getByRole('grid'));
    await settle();
    await press('F10');
    await press('v');
    h.ui.fireEvent.click(h.ui.getByRole('menuitemcheckbox', { name: 'Theme: dark' }));
    await settle();

    expect(picture().inputs.at(-1)).toBe(true);
    const { fills, strokes } = colours(picture().draw);
    // A dark panel with a light title, a border that shows, and gridlines that do not shout.
    expect(fills.slice(0, 2)).toEqual(['#1a1d23', '#f3f4f6']);
    expect(strokes).toEqual(['#4b515c', '#4b515c', '#2a2e36', '#2a2e36', '#2a2e36']);
    expect(fills).toContain('#9ca3af');
    expect(fills).not.toContain('surface');
    // The first series is the colour it was.
    expect(fills).toContain(SERIES_COLOURS[0]);
  });
});

describe('the ink a chart is drawn in', () => {
  it('is the theme\'s names in light and written-out colours in dark', () => {
    expect(LIGHT_INK).toEqual({ surface: 'surface', border: 'border', title: 'text', label: 'textMuted', grid: 'controlBorder', zero: 'border' });
    expect(DARK_INK).toEqual({ surface: '#1a1d23', border: '#4b515c', title: '#f3f4f6', label: '#9ca3af', grid: '#2a2e36', zero: '#4b515c' });
  });

  /** All eight read at three to one or better on the dark panel, so a series is the colour it was. */
  it('keeps the series colours in both themes', () => {
    expect(SERIES_COLOURS.map((_, index) => seriesColour(index, false))).toEqual(SERIES_COLOURS);
    expect(SERIES_COLOURS.map((_, index) => seriesColour(index, true))).toEqual(SERIES_COLOURS);
  });

  /** A colour that would not read is lifted a third of the way to white, keeping its hue. */
  it('lifts a colour too dark for the dark panel', () => {
    expect(readableOnDark('#1f3864')).toBe('#6a7a98');
    expect(readableOnDark('#3b6fd4')).toBe('#3b6fd4');
  });
});
