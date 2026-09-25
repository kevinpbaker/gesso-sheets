import { afterEach, describe, expect, it } from 'vitest';

import { createComponent, RouterService } from 'gesso-framework';
import { renderTest, serveForTest, type Rendered, type ServedForTest } from 'gesso-testing';
import 'gesso-testing/matchers';

import { isProofPath, PROOF_PATH } from '../route';
import { AppRoot, ROUTES, SHEET } from './Routes';
import { COMMANDS, menusFor, MENUS, SEPARATOR } from './SheetCommands';
import { SheetDocument } from './SheetDocument';
import { sheetChannel } from './sheetChannel';
import { SheetService } from './SheetService';

/**
 * The two routes, and the one thing that differs between them.
 *
 * The claim is not that a spreadsheet can be hidden behind a url. It
 * is that the default page is the application and nothing else: a
 * person who opens this project gets a spreadsheet, and the
 * instruments — the black strip, the heatmap, the button that
 * recalculates two hundred thousand cells — are on a second url,
 * where somebody who wants to be convinced can go and find them.
 *
 * The strip itself is DOM on the main thread and has no node in this
 * tree, so what is asserted here is the half the router owns: which
 * screen resolves, and what the chrome offers once it has.
 */
interface Harness {
  ui: Rendered;
  served: ServedForTest;
}

async function mount(): Promise<Harness> {
  const document = new SheetDocument();
  document.sheet.recalculate();
  const service = new SheetService(document, { rowCount: 200, columnCount: 20 });
  const served = serveForTest([sheetChannel(service)]);
  const ui = renderTest(createComponent(AppRoot), {
    channels: served.registry,
    routes: { routes: ROUTES, notFound: SHEET },
    width: 900,
    height: 420
  });
  await ui.settle();
  await served.settle();
  await ui.settle();
  return { ui, served };
}

describe('the two routes', () => {
  let h: Harness;

  afterEach(() => {
    h?.ui.unmount();
    h?.served.dispose();
  });

  /** Sends the app to a url the way the address bar would. */
  async function go(url: string): Promise<void> {
    h.ui.runtime.services.get(RouterService).navigate(url);
    await h.ui.settle();
    await h.served.settle();
    await h.ui.settle();
  }

  it('opens onto the sheet, with no recalculate button on it', async () => {
    h = await mount();
    expect(h.ui.queryByRole('grid')).not.toBeNull();
    expect(h.ui.queryByLabel(COMMANDS.recalculate.label)).toBeNull();
  });

  it('offers the button on the proof route', async () => {
    h = await mount();
    await go(PROOF_PATH);
    expect(h.ui.queryByRole('grid')).not.toBeNull();
    expect(h.ui.queryByLabel(COMMANDS.recalculate.label)).not.toBeNull();
  });

  /**
   * The same sheet on both, which is the point: the proof route is
   * not a second application, it is this one with instruments on it.
   */
  it('draws the same chrome either way', async () => {
    h = await mount();
    const plain = h.ui.getAllByRole('button').length;
    await go(PROOF_PATH);
    expect(h.ui.getAllByRole('button').length).toBe(plain + 1);
  });

  it('hands an unknown url the spreadsheet rather than an apology', async () => {
    h = await mount();
    await go('/nothing/here');
    expect(h.ui.queryByRole('grid')).not.toBeNull();
    expect(h.ui.queryByLabel(COMMANDS.recalculate.label)).toBeNull();
  });
});

/**
 * The shell reads the url for itself, because the strip is its DOM
 * and the router is in the worker. Two answers to one question is a
 * pair that can disagree, so the rule the router applies to a path is
 * the rule this applies too.
 */
describe('the path the shell matches', () => {
  it('is the proof route however it is spelled', () => {
    expect(isProofPath(PROOF_PATH)).toBe(true);
    expect(isProofPath(`${PROOF_PATH}/`)).toBe(true);
    expect(isProofPath('proof')).toBe(true);
  });

  it('is nothing else', () => {
    expect(isProofPath('/')).toBe(false);
    expect(isProofPath('')).toBe(false);
    expect(isProofPath('/proofs')).toBe(false);
    expect(isProofPath('/proof/deep')).toBe(false);
  });
});

describe('the menus a route shows', () => {
  it('keeps the whole bar on the proof route', () => {
    expect(menusFor(true)).toBe(MENUS);
  });

  it('drops the proof-only commands everywhere else', () => {
    const entries = menusFor(false).flatMap(menu => menu.entries);
    expect(entries).not.toContain('recalculate');
    // Every other command survives, so the filter is a filter and not
    // a second, shorter menu bar that will drift from the first.
    const all = MENUS.flatMap(menu => menu.entries).filter(entry => entry !== SEPARATOR);
    expect(entries.filter(entry => entry !== SEPARATOR)).toEqual(all.filter(entry => entry !== 'recalculate'));
  });

  /**
   * Data ends `SEPARATOR, 'recalculate'`, so the plain filter leaves a
   * rule drawn under nothing. The table itself is held to this by
   * `SheetCommands.spec`; what the table produces is held to it here.
   */
  it('leaves no menu opening or closing on a rule', () => {
    for (const menu of menusFor(false)) {
      expect(menu.entries[0], menu.id).not.toBe(SEPARATOR);
      expect(menu.entries[menu.entries.length - 1], menu.id).not.toBe(SEPARATOR);
      for (let index = 1; index < menu.entries.length; index++) {
        expect(menu.entries[index] === SEPARATOR && menu.entries[index - 1] === SEPARATOR, menu.id).toBe(false);
      }
    }
  });
});
