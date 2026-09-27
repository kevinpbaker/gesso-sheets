import { describe, expect, it } from 'vitest';

import { renderTest, serveForTest, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';
import { createComponent } from 'gesso-framework';
import type { UiKeyModifiers, UiNode } from 'gesso-core';

import { relativeRef } from '../sheet/A1';
import { SheetApp } from './SheetApp';
import { SheetDocument } from './SheetDocument';
import { sheetChannel } from './sheetChannel';
import { SheetService } from './SheetService';

/**
 * Clicking a chart opens the chart bar for it: what it reads, in the
 * colours the grid outlines it in, and a range to type.
 */

describe('clicking a chart', () => {
  let h: { ui: Rendered; served: ServedForTest; service: SheetService; document: SheetDocument };

  async function mount(): Promise<void> {
    const document = new SheetDocument();
    const rows = [
      ['Month', 'Units', 'Cost'],
      ['Jul', '5', '50'],
      ['Aug', '10', '90'],
      ['Sep', '15', '120']
    ];
    rows.forEach((line, row) => line.forEach((value, column) => document.setCell(row, column, value)));
    document.addSheet('Other');
    document.activate(0);
    document.addChart({
      kind: 'column',
      title: 'Units',
      legend: true,
      range: { start: relativeRef(0, 0), end: relativeRef(3, 2) },
      place: { x: 420, y: 20, width: 320, height: 200 }
    });
    document.sheet.recalculate();
    const service = new SheetService(document, { rowCount: 100, columnCount: 10 });
    // With no repository this publishes the document as it is, charts and all.
    await service.restore();
    const served = serveForTest([sheetChannel(service)]);
    const ui = renderTest(createComponent(SheetApp), { channels: served.registry, width: 1100, height: 700 });
    h = { ui, served, service, document };
    await settle();
  }

  async function settle(): Promise<void> {
    await h.ui.settle();
    await h.served.settle();
    await h.ui.settle();
  }

  async function press(key: string, modifiers: Partial<UiKeyModifiers> = {}): Promise<void> {
    h.ui.fireEvent.press(key, modifiers);
    await settle();
  }

  const bar = (): UiNode | null => h.ui.queryByRole('form', { name: 'Chart' });
  const says = (text: string | RegExp): boolean => {
    try {
      return h.ui.getAllByText(text).length > 0;
    } catch {
      return false;
    }
  };

  /** A real press and release at a node's centre, through the hit tester, as a pointer makes one. */
  async function pointAt(node: UiNode): Promise<void> {
    const box = h.ui.getLayout(node);
    h.ui.fireEvent.pointerDown(box.x + box.width / 2, box.y + box.height / 2);
    h.ui.fireEvent.pointerUp(box.x + box.width / 2, box.y + box.height / 2);
    await settle();
  }

  const clickChart = (): Promise<void> => pointAt(h.ui.getByRole('image', { name: /Units, chart of/ }));

  async function typeRange(text: string): Promise<void> {
    const field = h.ui.getByRole('textbox', { name: 'Chart range' });
    h.ui.fireEvent.focus(field);
    await press('a', { ctrl: true });
    h.ui.fireEvent.type(text);
    await press('Enter');
  }

  it('opens the chart bar, which says what the chart reads', async () => {
    await mount();
    expect(bar()).toBeNull();
    await clickChart();

    expect(bar()).not.toBeNull();
    expect(h.ui.getLayout(bar()!).height).toBeGreaterThan(20);
    const field = h.ui.getByRole('textbox', { name: 'Chart range' });
    expect(field).toHaveText('A1:C4');
    // Switched in when a chart is chosen, and laid out: not a field of no size with the right text in it.
    expect(h.ui.getLayout(field).width).toBeGreaterThan(100);
    expect(says('values B2:C4')).toBe(true);
    expect(says('labels A2:A4')).toBe(true);
    expect(says('names B1:C1')).toBe(true);
    expect(says('2 series, one per column.')).toBe(true);
  });

  it('takes a range typed into it, and says so when it cannot', async () => {
    await mount();
    await clickChart();

    await typeRange('A1:B4');
    expect(h.document.charts[0].range).toMatchObject({ start: { row: 0, column: 0 }, end: { row: 3, column: 1 } });
    expect(says('one series, from its column.')).toBe(true);

    await typeRange('Nowhere!A1:B2');
    expect(says('There is no sheet called Nowhere.')).toBe(true);
    expect(h.document.charts[0].range).toMatchObject({ end: { row: 3, column: 1 } });

    await typeRange('not a range');
    expect(says(/is not a range/)).toBe(true);

    await typeRange("'Other'!A1:B2");
    expect(h.document.charts[0].range.start.sheet).toBe('Other');
  });

  it('goes away when the chart is let go, but not a bar somebody opened', async () => {
    await mount();
    await clickChart();
    expect(bar()).not.toBeNull();
    await pointAt(h.ui.getByRole('cell', { name: 'Month' }));
    expect(bar()).toBeNull();

    // Opened from the menu, it stays open over a click on the grid.
    await press('F10');
    await press('i');
    h.ui.fireEvent.click(h.ui.getByRole('menuitem', { name: /Chart/ }));
    await settle();
    expect(bar()).not.toBeNull();
    await pointAt(h.ui.getByRole('cell', { name: 'Jul' }));
    expect(bar()).not.toBeNull();
  });
});
