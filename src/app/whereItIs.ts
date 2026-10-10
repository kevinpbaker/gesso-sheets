import type { SheetDocumentView } from './SheetContract';

/**
 * What the title bar says about where a workbook is, and whether it is
 * safe — Phase 36.
 *
 * A sentence rather than an icon, because the question it answers is
 * "where is my work?" and a cloud with a tick in it does not say
 * *this browser* or *Budget.gsheet*. Pure, and on the render worker's
 * side of the barrier, so the whole of what a person can be told is
 * one table a spec reads.
 */
export interface SaveSaid {
  readonly text: string;
  /** `danger` for anything that means the work is not being kept. */
  readonly tone: 'muted' | 'danger';
}

export function saveSaid(document: SheetDocumentView): SaveSaid {
  if (document.id === '') {
    return { text: '', tone: 'muted' };
  }
  if (document.elsewhere) {
    return { text: 'Open in another tab, so not saved here', tone: 'danger' };
  }
  switch (document.saving) {
    case 'failed':
      return { text: 'Not saved: this browser refused to store it', tone: 'danger' };
    case 'memory':
      return { text: 'Not saved: this browser keeps nothing for this site', tone: 'danger' };
    case 'saving':
      return { text: 'Saving…', tone: 'muted' };
    case 'off':
      return { text: '', tone: 'muted' };
  }
  const file = document.file;
  if (file === null) {
    return { text: 'Saved in this browser', tone: 'muted' };
  }
  if (document.autosave) {
    return { text: document.edited ? `Saving to ${file.name}…` : `Saved to ${file.name}`, tone: 'muted' };
  }
  if (file.handle === null) {
    return {
      text: document.edited ? `Saved in this browser · changed since ${file.name} was downloaded` : `Saved in this browser · downloaded as ${file.name}`,
      tone: 'muted'
    };
  }
  return {
    text: document.edited ? `Saved in this browser · ${file.name} is behind` : `Saved in this browser and to ${file.name}`,
    tone: 'muted'
  };
}

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

function clock(date: Date): string {
  return `${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}`;
}

/**
 * When something happened, as the home screen and the version list
 * say it: close times by distance, older ones by the calendar, in the
 * browser's own time zone.
 */
export function whenSaid(at: number, now: number): string {
  const ago = now - at;
  if (ago < 60_000) {
    return 'Just now';
  }
  if (ago < 60 * 60_000) {
    const minutes = Math.floor(ago / 60_000);
    return minutes === 1 ? '1 minute ago' : `${minutes} minutes ago`;
  }
  const then = new Date(at);
  const today = new Date(now);
  const yesterday = new Date(now);
  yesterday.setDate(today.getDate() - 1);
  const sameDay = (a: Date, b: Date) =>
    a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
  if (sameDay(then, today)) {
    return `Today, ${clock(then)}`;
  }
  if (sameDay(then, yesterday)) {
    return `Yesterday, ${clock(then)}`;
  }
  const day = `${then.getDate()} ${MONTHS[then.getMonth()]}`;
  return then.getFullYear() === today.getFullYear() ? `${day}, ${clock(then)}` : `${day} ${then.getFullYear()}`;
}

/** The workbooks whose name holds every word of a search, case aside. */
export function matching<T extends { readonly name: string }>(entries: readonly T[], search: string): T[] {
  const words = search.toLowerCase().split(/\s+/).filter(word => word !== '');
  return entries.filter(entry => words.every(word => entry.name.toLowerCase().includes(word)));
}
