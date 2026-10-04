import { afterEach, describe, expect, it } from 'vitest';

import { darkTheme, lightTheme, percent, UiEnvironmentKeys, type UiTheme } from 'gesso-core';
import { createComponent, ShellService, type ComponentContext, type Inputs } from 'gesso-framework';
import { renderTest, serveForTest, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';

import { appearanceTheme } from './appearance';
import { SheetApp } from './SheetApp';
import { SheetDocument } from './SheetDocument';
import { sheetChannel } from './sheetChannel';
import { SheetService } from './SheetService';

/**
 * View ▸ Appearance: light, dark, or the system's, and the menu says
 * which.
 *
 * Mounted under the same theme `AppRoot` gives the screen, rather than
 * through the router, because the theme is the part under test and the
 * router has nothing to do with it.
 */

interface Harness {
  ui: Rendered;
  served: ServedForTest;
}

let h: Harness | undefined;

afterEach(() => {
  h?.ui.unmount();
  h?.served.dispose();
  h = undefined;
});

function Themed(_inputs: Inputs<{}>, ctx: ComponentContext) {
  return (
    <column theme={appearanceTheme(ctx.inject(ShellService))} width={percent(100)} height={percent(100)}>
      {createComponent(SheetApp)}
    </column>
  );
}

async function settle(): Promise<void> {
  for (let round = 0; round < 3; round++) {
    await h!.ui.settle();
    await h!.served.settle();
  }
}

async function mount(): Promise<void> {
  const service = new SheetService(new SheetDocument(), { rowCount: 50, columnCount: 10 });
  const served = serveForTest([sheetChannel(service)]);
  const ui = renderTest(createComponent(Themed), { channels: served.registry, width: 1400, height: 600 });
  h = { ui, served };
  await settle();
}

function theme(): UiTheme {
  return h!.ui.getByRole('grid').environment!.get(UiEnvironmentKeys.theme);
}

async function openView(): Promise<void> {
  h!.ui.fireEvent.click(h!.ui.getByText('View'));
  await h!.ui.settle();
}

function ticked(): string[] {
  return [...h!.ui.semanticsTree().values()]
    .filter(record => record.role === 'menuitemcheckbox' && record.states?.includes('checked'))
    .map(record => record.label ?? '');
}

describe('View ▸ Appearance', () => {
  it('follows the system until something else is chosen, and says so', async () => {
    await mount();
    expect(theme()).toBe(lightTheme);

    h!.ui.runtime.setColorScheme('dark');
    await settle();
    expect(theme()).toBe(darkTheme);

    await openView();
    expect(ticked()).toEqual(['Match system appearance']);
  });

  it('holds a chosen appearance whatever the system reports', async () => {
    await mount();
    await openView();
    h!.ui.fireEvent.click(h!.ui.getByRole('menuitemcheckbox', { name: 'Dark appearance' }));
    await settle();
    expect(theme()).toBe(darkTheme);

    h!.ui.runtime.setColorScheme('light');
    await settle();
    expect(theme()).toBe(darkTheme);

    await openView();
    expect(ticked()).toEqual(['Dark appearance']);
    h!.ui.fireEvent.click(h!.ui.getByRole('menuitemcheckbox', { name: 'Match system appearance' }));
    await settle();
    expect(theme()).toBe(lightTheme);
  });
});
