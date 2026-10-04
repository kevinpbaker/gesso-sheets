import { combineLatest, distinctUntilChanged, map, type Observable } from 'rxjs';
import { darkTheme, lightTheme, type UiTheme } from 'gesso-core';
import { persisted, ShellStorage, type ColorScheme, type PersistedState, type ShellService } from 'gesso-framework';

/**
 * Light, dark, or whatever the system says: View ▸ Appearance.
 *
 * `system` is kept as itself and not resolved to light or dark when it
 * is chosen, because the platform can change its mind while the sheet
 * is open, and following it an hour later is the point of choosing it.
 */
export type Appearance = 'system' | ColorScheme;

const APPEARANCES: readonly Appearance[] = ['system', 'light', 'dark'];

/**
 * One choice per shell, shared by whoever asks.
 *
 * Two readers need the same cell — the root, which themes everything
 * below it, and the menu, which shows and changes it — and neither is
 * inside the other's props, with the router between them. A second
 * `persisted` on the same key would be a second cell that a choice in
 * the menu never reaches, so there is one, made on first use.
 *
 * Kept like the status bar's figures: in the browser's storage,
 * through the shell, because it is the person's and not the
 * document's, and wanted before the first frame of the next visit.
 */
const chosen = new WeakMap<ShellService, PersistedState<Appearance>>();

export function appearance(shell: ShellService): PersistedState<Appearance> {
  let state = chosen.get(shell);
  if (state === undefined) {
    state = persisted(new ShellStorage(shell, { prefix: 'gessosheet:' }), 'appearance', {
      initial: 'system',
      revive: raw => (APPEARANCES.includes(raw as Appearance) ? (raw as Appearance) : null),
      label: 'appearance'
    });
    chosen.set(shell, state);
  }
  return state;
}

/** The theme the choice comes to, given what the platform reports. */
export function appearanceTheme(shell: ShellService): Observable<UiTheme> {
  return combineLatest([appearance(shell).value, shell.colorScheme]).pipe(
    map(([choice, reported]) => (choice === 'system' ? reported : choice)),
    distinctUntilChanged(),
    map(scheme => (scheme === 'dark' ? darkTheme : lightTheme))
  );
}
