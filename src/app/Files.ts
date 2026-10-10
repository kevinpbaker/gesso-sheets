import type { Observable } from 'rxjs';

import {
  internalState,
  RouterService,
  type ChannelReplica,
  ShellService,
  type ComponentContext,
  type ShellFileResult,
  type ShellFileType,
  type ShellRecentFile
} from 'gesso-framework';

import { base64OfBytes, bytesOfBase64 } from './base64';
import type { SheetCommands, SheetView } from './SheetContract';

/**
 * Files, on the render worker: the route, the pickers, and the saves.
 *
 * Three threads take part in every one of these, and this is the one
 * in the middle. The application worker knows what a workbook *is* and
 * builds it; the shell has the pickers and the handles; this thread
 * has the click that a picker needs, and the route that says which
 * document a tab is. So it carries each request from the one that
 * knows to the one that can.
 *
 * ## A tab is a document
 *
 * The route names it: `/d/<id>`. `/` means "whichever was used last",
 * and `/d/new` a blank one, and in both cases the url is replaced with
 * the real id as soon as the application worker says what it opened —
 * so a reload reopens that document, and not the last-used one or
 * another blank. Opening a file *pushes* instead, so Back is the
 * document that was there before.
 */

export const WORKBOOK: ShellFileType = {
  description: 'Gessosheet workbook',
  mediaType: 'application/json',
  extensions: ['.gsheet']
};

export const XLSX: ShellFileType = {
  description: 'Excel workbook',
  mediaType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  extensions: ['.xlsx']
};

export const CSV: ShellFileType = {
  description: 'Comma-separated values',
  mediaType: 'text/csv',
  extensions: ['.csv', '.tsv', '.txt']
};

export interface FileActions {
  /** A blank document, in a tab of its own. */
  newDocument(): void;
  /** An open picker: a workbook opens as the document here, a CSV as a sheet. */
  open(): void;
  /** A workbook from the library, by id, in this tab. */
  openDocument(id: string): void;
  /** A picker for a CSV or an Excel workbook: File ▸ Import. */
  importFile(): void;
  /** The print window, with the sheet in view on it: File ▸ Print, or as a PDF. */
  print(pdf: boolean): void;
  /** One of the files the shell remembers. */
  reopen(handle: number): void;
  /** Save, or Save As. */
  save(asNew: boolean): void;
  /** The sheet in view, as a CSV. */
  exportCsv(): void;
  /** The workbook, as an Excel file. */
  exportXlsx(): void;
  /** What the shell remembers, refreshed by `refreshRecent`. */
  readonly recent: Observable<readonly ShellRecentFile[]> & { readonly value: readonly ShellRecentFile[] };
  refreshRecent(): void;
}

export function fileActions(ctx: ComponentContext, sheet: ChannelReplica<SheetView, SheetCommands>): FileActions {
  const shell = ctx.inject(ShellService);
  const router = ctx.inject(RouterService);
  const recent = internalState<readonly ShellRecentFile[]>([]);

  /** Set by an open, so the url the document lands on is pushed rather than replaced. */
  let push = false;

  // Which document the route names, sent whenever the route moves.
  // Sending the one already open is free on the other side, so this
  // does not have to know whether the move was its own.
  ctx.effect(router.match, match => {
    if (match === null || match.path === '/proof') {
      if (match?.path === '/proof') {
        sheet.send.openDocument('');
      }
      return;
    }
    sheet.send.openDocument(match.params.id ?? '');
  });

  // And the url brought into line with what opened.
  ctx.effect(sheet.view.document, document => {
    const match = router.match.value;
    if (document.id === '' || match === null || match.path === '/proof') {
      return;
    }
    const wanted = `/d/${document.id}`;
    if (match.path !== wanted) {
      const named = match.params.id;
      router.navigate(wanted, { replace: !push || named === undefined || named === 'new' });
    }
    push = false;
  });

  /**
   * A download the application worker built, handed to the shell.
   *
   * Starting from the serial already published, not from zero: this
   * is set up again whenever the screen is, and a download that was
   * answered before must not be saved a second time because the thread
   * asking has forgotten it asked.
   */
  let handled = sheet.view.transfer.value.download?.serial ?? 0;
  ctx.effect(sheet.view.transfer, transfer => {
    const download = transfer.download;
    if (download === null || download.serial <= handled) {
      return;
    }
    handled = download.serial;
    void shell
      .saveFile({
        name: download.name,
        text: download.text,
        ...(download.base64 === undefined ? {} : { bytes: bytesOfBase64(download.base64) }),
        mediaType: download.mediaType,
        accept: [download.kind === 'workbook' ? WORKBOOK : download.kind === 'xlsx' ? XLSX : CSV],
        ...(download.handle === null ? {} : { handle: download.handle })
      })
      .then(result => {
        if (result.outcome === 'ok' && result.saved !== null) {
          sheet.send.fileSaved(download.kind, result.saved.name, result.saved.handle, result.saved.via, download.quiet === true);
        } else if (download.quiet === true) {
          sheet.send.fileNotSaved(reason(result));
        } else if (result.outcome !== 'cancelled') {
          sheet.send.reportFile(`${download.name} was not saved: ${reason(result)}`);
        }
      });
  });

  const deliver = (result: ShellFileResult): void => {
    if (result.outcome === 'cancelled') {
      return;
    }
    if (result.outcome !== 'ok') {
      sheet.send.reportFile(`Nothing was opened: ${reason(result)}`);
      return;
    }
    for (const file of result.files) {
      push ||= isWorkbook(file.name);
      sendFile(sheet, file.name, file.bytes, file.handle);
    }
  };

  return {
    /**
     * A blank workbook, here, with Back to return to this one.
     *
     * In this tab rather than a new one, which is what it did until
     * Phase 36: a new tab is a second application worker starting from
     * nothing, a second or two of blank page, and a popup blocker's to
     * refuse — in a headless browser it simply did not arrive.
     *
     * Asked of the application worker directly, and the url follows
     * what opens, pushed so Back returns: the path a file takes.
     */
    newDocument: () => {
      push = true;
      sheet.send.openDocument('new');
    },
    openDocument: id => {
      push = true;
      sheet.send.openDocument(id);
    },
    open: () => void shell.openFiles({ accept: [WORKBOOK, XLSX, CSV], multiple: true }).then(deliver),
    reopen: handle => void shell.reopenFile(handle).then(deliver),
    importFile: () => void shell.openFiles({ accept: [CSV, XLSX], multiple: true }).then(deliver),
    /**
     * The window first, while the click that asked is fresh — a popup
     * asked for after a round trip to the other worker is one the
     * browser may already have stopped allowing — and the page second,
     * which the application worker builds and sends it.
     */
    print: pdf => {
      void shell.openPopup({ url: '/print.html', name: 'gessosheet-print', width: 1000, height: 760 }).then(opened => {
        if (!opened) {
          sheet.send.reportFile('The print window was blocked. Allow pop-ups for this site, then print again.');
        }
      });
      sheet.send.print(pdf);
    },
    save: asNew => sheet.send.saveDocument(asNew),
    exportCsv: () => sheet.send.exportCsv(),
    exportXlsx: () => sheet.send.exportXlsx(),
    recent,
    refreshRecent: () =>
      void shell.recentFiles().then(result => {
        recent.value = result.outcome === 'ok' ? result.recent : [];
      })
  };
}

/** Whether a file opens as a document of its own, rather than as a sheet in this one. */
function isWorkbook(name: string): boolean {
  return /\.(gsheet|json|xlsx)$/i.test(name);
}

/**
 * A file's bytes, sent to whichever command reads its kind.
 *
 * An `.xlsx` crosses as its bytes — it is a zip, and decoding it as
 * text here would destroy it — and everything else as UTF-8 text,
 * decoded once on this side, which is where the file arrived. For a
 * drop and a pick alike: the two differ only in whether there is a
 * handle to remember.
 */
export function sendFile(
  sheet: ChannelReplica<SheetView, SheetCommands>,
  name: string,
  bytes: ArrayBuffer | undefined,
  handle: number | null
): void {
  const raw = new Uint8Array(bytes ?? new ArrayBuffer(0));
  if (/\.xlsx$/i.test(name)) {
    sheet.send.importXlsx(name, base64OfBytes(raw));
    return;
  }
  sheet.send.openFile(name, new TextDecoder().decode(raw), handle);
}

/** Why, as the end of a sentence. */
function reason(result: ShellFileResult): string {
  switch (result.outcome) {
    case 'denied':
      return 'the browser did not allow it.';
    case 'unsupported':
      return 'this browser cannot reach files.';
    default:
      return result.error === null ? 'something went wrong.' : `${result.error.replace(/\.$/, '')}.`;
  }
}
