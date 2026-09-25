/**
 * Which page is being asked for, for the one thread that cannot ask
 * the router.
 *
 * There are two urls. `/` is the spreadsheet, and it is the whole of
 * what a person who wants a spreadsheet should be given. `/proof` is
 * the same spreadsheet with the instruments attached: the black strip
 * along the top, the main-thread pulse, the block button, the layout
 * heatmap, and the one command that recalculates two hundred thousand
 * cells so that there is something for the strip to be measuring.
 *
 * The routes themselves are declared in `src/app/Routes.tsx`, in the
 * render worker, because a route names a component class and no class
 * crosses a `postMessage`. What cannot be declared there is the half
 * of the proof route that is *not* in the worker — the strip is
 * ordinary DOM on the main thread, which is the entire reason it is
 * believable — so the shell has to answer the same question for
 * itself, from `location`. This file is the answer both of them use,
 * and it is deliberately the smallest module in the project: it is
 * reachable from the main thread as well as from the worker, which is
 * the one thing that turns a save into a page reload.
 */

/** Where the proof route lives. */
export const PROOF_PATH = '/proof';

/**
 * Whether a url is the proof route.
 *
 * Leading, trailing and doubled slashes are ignored, because the
 * router ignores them: `/proof/` resolves to the proof screen, and a
 * shell that disagreed would draw the strip over the plain sheet or
 * the plain page over the proof one.
 */
export function isProofPath(pathname: string): boolean {
  return segments(pathname).join('/') === segments(PROOF_PATH).join('/');
}

function segments(path: string): string[] {
  return path.split('/').filter(segment => segment.length > 0);
}
