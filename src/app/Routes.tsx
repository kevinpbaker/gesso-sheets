import { percent } from 'gesso-core';
import { createComponent, route, RouterOutlet, type ComponentContext, type Inputs } from 'gesso-framework';

import { PROOF_PATH } from '../route';
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
export const PROOF = route({ path: PROOF_PATH, component: ProofScreen });

export const ROUTES = [SHEET, PROOF];

/**
 * The application's root, which is an outlet and a background.
 *
 * `RouterOutlet` cannot *be* the root — its render returns an
 * Observable, and the layout root has to be a box — so it goes inside
 * one, which is where it wanted to be anyway. Nothing else is here on
 * purpose: chrome drawn around the outlet would be chrome the proof
 * route and the plain route shared, and they share the screen rather
 * than a frame around it.
 */
export function AppRoot(_inputs: Inputs<{}>, _ctx: ComponentContext) {
  return (
    <column width={percent(100)} height={percent(100)} backgroundColor="background">
      {createComponent(RouterOutlet)}
    </column>
  );
}
