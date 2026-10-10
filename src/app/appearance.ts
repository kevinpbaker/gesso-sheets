import { combineLatest, map, type Observable } from 'rxjs';

import { darkTheme, lightTheme, type UiTheme } from 'gesso-core';
import { persisted, ShellService, ShellStorage, type ComponentContext, type PersistedState } from 'gesso-framework';

/**
 * Light, dark, or whatever the platform says.
 *
 * `auto` is the default because it is the only choice that asks the
 * person nothing: a dark desktop gets a dark sheet. The other two are
 * for the person whose desktop and spreadsheet disagree, which is
 * common enough that every spreadsheet with a dark mode has the
 * override.
 */
export type Appearance = 'auto' | 'light' | 'dark';

export const APPEARANCES: readonly Appearance[] = ['auto', 'light', 'dark'];

function isAppearance(raw: unknown): raw is Appearance {
  return typeof raw === 'string' && (APPEARANCES as readonly string[]).includes(raw);
}

/**
 * One preference per shell, made by whichever component asks first.
 *
 * The root reads it to pick the theme and the menu bar writes it. A
 * second `persisted` on the same key would be a second copy that the
 * first one's writes never reach, and the specs mount the menu bar
 * without the root, so neither can own it. Keyed by the shell rather
 * than held once for the module because a spec process runs many
 * applications, each with a shell and a storage of its own. It lives
 * as long as its shell does, which is the life of the page.
 */
const preferences = new WeakMap<ShellService, PersistedState<Appearance>>();

export function appearancePreference(ctx: ComponentContext): PersistedState<Appearance> {
  const shell = ctx.inject(ShellService);
  let preference = preferences.get(shell);
  if (preference === undefined) {
    preference = persisted(new ShellStorage(shell, { prefix: 'gessosheet:' }), 'appearance', {
      initial: 'auto',
      revive: raw => (isAppearance(raw) ? raw : null),
      label: 'appearance'
    });
    preferences.set(shell, preference);
  }
  return preference;
}

/** The theme to paint with: the person's choice, or the platform's when they made none. */
export function appearanceTheme(ctx: ComponentContext): Observable<UiTheme> {
  const shell = ctx.inject(ShellService);
  return combineLatest([appearancePreference(ctx).value, shell.colorScheme]).pipe(
    map(([chosen, platform]) => ((chosen === 'auto' ? platform : chosen) === 'dark' ? darkTheme : lightTheme))
  );
}
