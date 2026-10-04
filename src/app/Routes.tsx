import { percent } from 'gesso-core';
import { createComponent, route, RouterOutlet, ShellService, type ComponentContext, type Inputs } from 'gesso-framework';

import { PROOF_PATH } from '../route';
import { appearanceTheme } from './appearance';
import { SheetApp } from './SheetApp';

/**
 * The two pages, and the root that swaps between them.
 *
 * They are declared here, in the render worker, because a route names
 * a component and a component cannot cross a `postMessage`. The shell
 * holds no routes at all: it reports the url the window is at and
 * performs the pushes this side asks for, and that is the whole of
 * its half of routing. `src/route.ts` is the one fact both threads
 * need — the path itself — and the reason the shell needs it is that
 * the proof route's strip is DOM, which is not this thread's to draw.
 *
 * Two screens rather than one screen reading the url, because a route
 * component takes no props: the difference between the pages is a
 * boolean, and these two lines are where it is supplied. `SheetApp`
 * is the same component on both, mounted fresh on each — a navigation
 * between them rebuilds the tree, and the sheet itself does not
 * notice, because the sheet is on the other thread and this one has
 * never held a cell.
 */
export function SheetScreen(_inputs: Inputs<{}>, _ctx: ComponentContext) {
  return <SheetApp />;
}

export function ProofScreen(_inputs: Inputs<{}>, _ctx: ComponentContext) {
  return <SheetApp proof={true} />;
}

export const SHEET = route({ path: '/', component: SheetScreen });
/**
 * One document, by the library's id for it. The same screen as `/`,
 * which is the last document used: which one to show is a thing
 * `Files.ts` reads from the route and tells the application worker,
 * so the screen itself does not change between them.
 */
export const DOCUMENT = route({ path: '/d/:id', component: SheetScreen });
export const PROOF = route({ path: PROOF_PATH, component: ProofScreen });

export const ROUTES = [SHEET, DOCUMENT, PROOF];

/**
 * The application's root, which is an outlet and a background.
 *
 * `RouterOutlet` cannot *be* the root — its render returns an
 * Observable, and the layout root has to be a box — so it goes inside
 * one, which is where it wanted to be anyway. Nothing else is here on
 * purpose: chrome drawn around the outlet would be chrome the proof
 * route and the plain route shared, and they share the screen rather
 * than a frame around it.
 *
 * The theme is the one thing set here, because it is the one thing
 * every screen inherits: light or dark as View ▸ Appearance says,
 * which by default is as the platform reports it.
 * Every role colour below — `background`, `text`, `border`, the
 * selection's wash — follows from it; the colours a person chose for
 * a cell are their data and do not.
 */
export function AppRoot(_inputs: Inputs<{}>, ctx: ComponentContext) {
  const theme = appearanceTheme(ctx.inject(ShellService));
  return (
    <column theme={theme} width={percent(100)} height={percent(100)} backgroundColor="background">
      {createComponent(RouterOutlet)}
    </column>
  );
}
