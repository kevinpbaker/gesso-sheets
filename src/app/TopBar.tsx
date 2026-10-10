import { BehaviorSubject, combineLatest, distinctUntilChanged, map, type Observable } from 'rxjs';

import {
  editorFor,
  percent,
  type UiKeyboardEvent,
  type UiNode,
  type UiSemanticState,
  type UiTextChangeEvent,
  type UiTextSpan
} from 'gesso-core';
import { ColorPalette, ColorPicker, Dialog } from 'gesso-components';
import { FocusService, internalState, type ComponentContext, type Inputs } from 'gesso-framework';

import { appearancePreference } from './appearance';
import { FindBar } from './FindBar';
import { formulaSpans } from './FormulaColours';
import { MenuBar } from './MenuBar';
import { NameBox, ONE_CELL } from './NameBox';
import { ChartBar } from './ChartBar';
import type { FileActions } from './Files';
import { RecentBar } from './RecentBar';
import { RulesBar, type RulesTab } from './RulesBar';
import { TAB_COLOURS } from './SheetTabs';
import { cornerOf, isOneCell, Sheet, zoomStep } from './SheetContract';
import {
  acceleratorLabel,
  CHART_POINTS,
  COMMANDS,
  FUNCTION_CALLS,
  menusFor,
  offers,
  STRESS_CELLS,
  type CommandId
} from './SheetCommands';
import type { SheetFormatChange } from './SheetContract';
import { columnName, formatRange, relativeRef } from '../sheet/A1';
import type { CellPaint } from '../sheet/Format';
import { ICONS } from './icons';
import { Toolbar, type ToolbarItem } from './Toolbar';
import type { SheetEditing } from './SheetEditing';
import { keyAction } from './SheetKeys';
import { colourName } from './colourNames';
import { NamesDialog } from './NamesDialog';
import { NoteDialog } from './NoteDialog';
import { ScriptDialog } from './ScriptDialog';
import { PasteHint, Shortcuts } from './Shortcuts';

/**
 * Everything above the grid: the menu bar, the toolbar, the name box
 * and the formula bar, and the find bar when it is open.
 *
 * All of it painted in the render worker, which is the part worth
 * saying out loud. The grid was one role repeated a thousand times;
 * this is fifteen roles with a traversal model, an overlay, a dialog
 * and four tab stops, and if a canvas UI cannot do *this* then it
 * cannot do applications — a fast grid with a DOM toolbar bolted on
 * top would be a demonstration rather than a program.
 *
 * It also owns command dispatch, because a command is the one thing
 * the menu bar, the toolbar and the accelerators all need and none of
 * them should each have their own copy of.
 */
export interface TopBarProps {
  readonly editing: SheetEditing;
  /** Open, save and the rest; see `Files.ts`. */
  readonly files: FileActions;
  /**
   * Whether this is the proof route, which is the only thing above
   * the grid that differs between the two.
   *
   * It reaches here as a prop rather than being read from a url,
   * because this thread has no url to read: the render worker's
   * `location` is the worker script's. The route is resolved by the
   * router a few files up and arrives as a boolean, which is all the
   * chrome ever needed to know.
   */
  readonly proof?: boolean;
}

/** Which half of the find bar is showing, or neither. */
type Finding = 'closed' | 'find' | 'replace';

export function TopBar(inputs: Inputs<TopBarProps>, ctx: ComponentContext) {
  const sheet = ctx.channel(Sheet);
  const focus = ctx.inject(FocusService);
  const edit = inputs.editing.value;
  const status = sheet.view.status;
  const appearance = appearancePreference(ctx);
  const proof = inputs.proof.value === true;
  const files = inputs.files.value;

  /** A property of the active cell's paint, as something to bind. */
  const on = (read: (paint: CellPaint) => boolean): Observable<boolean> =>
    sheet.view.activeFormat.pipe(map(current => read(current.paint)));

  /**
   * What a toolbar button is called, what its tooltip says, and what
   * pressing it does — all three from the command table.
   *
   * Every button on this row is one command, so none of the three is
   * worth writing out beside the icon: a label typed here is a label
   * that can drift from the menu's, and a tooltip is the one place in
   * this application where somebody finds out that Bold has a
   * shortcut. The accelerator comes from the same entry the menu
   * prints it from, so it is right or they are both wrong.
   */
  const runs = (id: CommandId): { label: string; tip: string; onRun: () => void } => {
    const accelerator = COMMANDS[id].accelerator;
    const label = COMMANDS[id].label;
    return {
      label,
      tip: accelerator === undefined ? label : `${label} (${acceleratorLabel(accelerator)})`,
      onRun: () => run(id)
    };
  };

  /**
   * Undo and Redo say what they would do — "Undo sort" — because the
   * history knows and a bare "Undo" makes somebody press it to find
   * out. Read when a menu opens or a tooltip shows, off the replica,
   * so it costs a property read and no subscription.
   */
  const labelNow = (id: CommandId): string | undefined => {
    if (id === 'showFormulas') {
      return status.value.showingFormulas ? 'Show values' : undefined;
    }
    if (id === 'showReferences') {
      return edit.referencesShown.value ? 'Hide references' : undefined;
    }
    if (id !== 'undo' && id !== 'redo') {
      return undefined;
    }
    const what = id === 'undo' ? status.value.undoLabel : status.value.redoLabel;
    return what === '' ? undefined : `${COMMANDS[id].label} ${what}`;
  };

  const tipFor = (id: 'undo' | 'redo'): string => {
    const accelerator = COMMANDS[id].accelerator;
    const label = labelNow(id) ?? COMMANDS[id].label;
    return accelerator === undefined ? label : `${label} (${acceleratorLabel(accelerator)})`;
  };

  /**
   * The toolbar, as data.
   *
   * A list rather than a row of elements so the toolbar can own its
   * own keyboard — see `Toolbar.tsx` — and so that what each button
   * does is the same `run` the menu and the accelerators go through.
   * Three ways to reach a command and one place that performs it.
   */
  /**
   * The text and fill palettes: whether each is open, and the button it
   * opens beside. The colours used lately are shared between the two, as
   * a spreadsheet's are, most recent first.
   */
  const textPaletteOpen = internalState(false);
  const fillPaletteOpen = internalState(false);
  const textAnchor = internalState<UiNode | null>(null);
  const fillAnchor = internalState<UiNode | null>(null);
  const recentColours = internalState<readonly string[]>([]);
  /**
   * The custom colour dialog: which of the two it is for, while it is
   * open, and the colour as it is dragged — applied only when it is kept.
   */
  const customFor = internalState<'color' | 'fill' | null>(null);
  const customColour = internalState('#000000');
  const openCustom = (which: 'color' | 'fill'): void => {
    const now = which === 'color' ? paint().color : paint().fill;
    customColour.value = now === '' ? (which === 'color' ? '#000000' : '#ffffff') : now;
    customFor.value = which;
  };
  const closeCustom = (): void => {
    customFor.value = null;
    edit.focusSheet();
  };
  const paintWith = (change: { color: string } | { fill: string }): void => {
    const colour = 'color' in change ? change.color : change.fill;
    if (colour !== '') {
      recentColours.value = [colour, ...recentColours.value.filter(each => each !== colour)].slice(0, 10);
    }
    sheet.send.format(change);
    edit.focusSheet();
  };

  const tools: readonly ToolbarItem[] = [
    {
      id: 'undo',
      icon: ICONS.undo,
      ...runs('undo'),
      tip: () => tipFor('undo'),
      enabled: status.pipe(map(s => s.canUndo))
    },
    {
      id: 'redo',
      icon: ICONS.redo,
      ...runs('redo'),
      tip: () => tipFor('redo'),
      enabled: status.pipe(map(s => s.canRedo))
    },
    { id: 'bold', icon: ICONS.bold, ...runs('bold'), startsGroup: true, pressed: on(p => p.bold) },
    { id: 'italic', icon: ICONS.italic, ...runs('italic'), pressed: on(p => p.italic) },
    { id: 'underline', icon: ICONS.underline, ...runs('underline'), pressed: on(p => p.underline) },
    {
      id: 'textColour',
      text: 'A',
      weight: 'bold',
      label: 'Text colour',
      onRun: () => run('textColour'),
      tip: () => `Text colour (${paint().color === '' ? 'automatic' : colourName(paint().color)})`,
      swatch: sheet.view.activeFormat.pipe(map(current => current.paint.color || 'text')),
      anchor: node => (textAnchor.value = node)
    },
    {
      id: 'fillColour',
      icon: ICONS.fill,
      label: 'Fill colour',
      onRun: () => run('fillColour'),
      tip: () => `Fill colour (${paint().fill === '' ? 'none' : colourName(paint().fill)})`,
      swatch: sheet.view.activeFormat.pipe(map(current => current.paint.fill)),
      anchor: node => (fillAnchor.value = node)
    },
    {
      id: 'alignLeft',
      icon: ICONS.alignLeft,
      ...runs('alignLeft'),
      startsGroup: true,
      pressed: on(p => p.align === 'start')
    },
    { id: 'alignCenter', icon: ICONS.alignCenter, ...runs('alignCenter'), pressed: on(p => p.align === 'center') },
    { id: 'alignRight', icon: ICONS.alignRight, ...runs('alignRight'), pressed: on(p => p.align === 'end') },
    {
      id: 'formatPainter',
      icon: ICONS.painter,
      ...runs('formatPainter'),
      pressed: edit.painter.pipe(map(state => state !== 'off')),
      // Two clicks keep it lit until Escape, for painting several places.
      onDoubleRun: () => {
        sheet.send.pickFormats();
        edit.setPainter('held');
      },
      tip: 'Format painter (double-click to keep it on)'
    },
    { id: 'wrap', text: 'Wrap', ...runs('wrap'), pressed: on(p => p.wrap) },
    { id: 'currency', icon: ICONS.currency, ...runs('formatCurrency'), startsGroup: true },
    { id: 'percent', icon: ICONS.percent, ...runs('formatPercent') },
    { id: 'fewerDecimals', text: '.0\u2190', ...runs('fewerDecimals') },
    { id: 'moreDecimals', text: '.00\u2192', ...runs('moreDecimals') },
    // The proof route's one extra button. A toolbar id is not a
    // command id — `currency` is `formatCurrency` here — so the
    // question is asked of the command rather than filtered out of
    // the row afterwards.
    ...(offers('recalculate', proof)
      ? [{ id: 'recalculate', icon: ICONS.recalculate, ...runs('recalculate'), startsGroup: true }]
      : [])
  ];

  const finding = internalState<Finding>('closed');
  /**
   * Whether the rules bar is open, and which half it is showing.
   *
   * Two states rather than one, and not for tidiness: written as one,
   * the bar's own `tab` prop was derived from the state that switches
   * the bar in and out — so swapping it in emitted into its own
   * props in the same turn, and the subtree arrived in the
   * accessibility tree having never been laid out. A browser showed
   * a bar of zero height with every role in it correct.
   */
  const ruling = internalState(false);
  const rulesTab = internalState<RulesTab>('format');
  /** Whether the chart bar is open; see `ruling` for why it is its own. */
  const charting = internalState(false);
  /**
   * Whether it was a click on a chart that opened the bar, rather than
   * Insert ▸ Chart. Clicking a chart shows what it reads and lets it be
   * changed, as Excel shows its chart tab; letting the chart go puts
   * away a bar that it opened, and leaves one somebody asked for.
   */
  let chartingForSelection = false;
  ctx.effect(
    sheet.view.charts.pipe(
      map(charts => charts.selected),
      distinctUntilChanged()
    ),
    selected => {
      if (selected !== 0 && !charting.value) {
        chartingForSelection = true;
        charting.value = true;
      } else if (selected === 0 && charting.value && chartingForSelection) {
        chartingForSelection = false;
        charting.value = false;
      }
    }
  );
  /** Whether the recent-files bar is open; see `ruling` for why it is its own. */
  const recenting = internalState(false);
  const shortcutsOpen = internalState(false);
  /**
   * The line under the bar: what the name box is waiting for, or why
   * it would not take what it was given.
   *
   * One slot rather than two, because it is one conversation. `Insert
   * ▸ Name` opens it by saying which range is about to be named, the
   * box answers in the same place when the name will not do, and a
   * keystroke that changes the name clears it.
   */
  const notice = internalState('');
  /** Where the paste hint is, when somebody asks for Paste from a menu. */
  const pasteHint = internalState(false);
  /**
   * The note being written, and the cell it is for — held from the
   * moment the dialog opens, so a note is saved on the cell it was
   * begun on even if the selection moves underneath it.
   */
  const noteOpen = internalState(false);
  const noteCell = internalState('');
  const noteText = internalState('');
  let noteAt = { row: 0, column: 0 };
  const closeNote = (): void => {
    noteOpen.value = false;
    edit.focusSheet();
  };
  /** Whether Insert ▸ Names is open. */
  const namesOpen = internalState(false);
  const closeNames = (): void => {
    namesOpen.value = false;
    edit.focusSheet();
  };
  /** Whether Data ▸ Scripts is open. */
  const scriptsOpen = internalState(false);
  const closeScripts = (): void => {
    scriptsOpen.value = false;
    edit.focusSheet();
  };

  let nameBoxNode: UiNode | null = null;
  let findFieldNode: UiNode | null = null;
  let menuBarNode: UiNode | null = null;
  /**
   * A field that has been asked for the keyboard but does not exist
   * yet.
   *
   * Ctrl+F mounts the find bar and wants its field focused, and the
   * field is not a node until the frame after. Rather than guess at
   * how long that takes, the request is remembered and the `ref`
   * callback honours it the moment the node arrives — which is the
   * exact moment, rather than a plausible one.
   */
  let rulesFieldNode: UiNode | null = null;
  let chartFieldNode: UiNode | null = null;
  let recentFieldNode: UiNode | null = null;
  let wanted: 'name' | 'find' | 'rules' | 'chart' | 'recent' | null = null;

  /**
   * Focus a field and select what is in it.
   *
   * Selecting is the half that is easy to leave out and impossible to
   * use without. The name box already holds `A1`, so a Ctrl+G that
   * only focused it put the caret at the front and `C120` typed into
   * it became `C120A1` — which is not an address, so the jump
   * silently did nothing. Typing over a focused field replaces it,
   * for the same reason typing over a selected cell does.
   */
  const take = (node: UiNode): void => {
    focus.focus(node);
    const model = editorFor(node);
    if (model !== null) {
      model.select(0, model.text.length);
    }
  };

  /**
   * Asks a field for the keyboard, mounting it first if it is not
   * there yet.
   *
   * The request is registered *before* `open` runs, and that order is
   * the whole of it. Written the other way round — flip the state,
   * then ask — the find bar mounted synchronously, its `ref` fired
   * while nothing had been asked for, and the request was left
   * waiting for a node that had already arrived. Ctrl+F opened a bar
   * and left the person typing into the sheet behind it.
   */
  const askFor = (which: 'name' | 'find' | 'rules' | 'chart' | 'recent', open?: () => void): void => {
    wanted = which;
    open?.();
    const node =
      which === 'name'
        ? nameBoxNode
        : which === 'find'
          ? findFieldNode
          : which === 'rules'
            ? rulesFieldNode
            : which === 'chart'
              ? chartFieldNode
              : recentFieldNode;
    if (wanted === which && node !== null) {
      wanted = null;
      take(node);
    }
  };

  const arrived = (which: 'name' | 'find' | 'rules' | 'chart' | 'recent') => (node: UiNode | null) => {
    if (which === 'name') {
      nameBoxNode = node;
    } else if (which === 'find') {
      findFieldNode = node;
    } else if (which === 'rules') {
      rulesFieldNode = node;
    } else if (which === 'chart') {
      chartFieldNode = node;
    } else {
      recentFieldNode = node;
    }
    if (node !== null && wanted === which) {
      wanted = null;
      take(node);
    }
  };

  /**
   * Whether a command can be run right now.
   *
   * Read synchronously off the replica rather than piped, because the
   * menu asks at the moment it opens and an Observable would be a
   * subscription per item per open. The values are already here — a
   * replica's view keys are cells with a current value — so the
   * question costs a property read.
   */
  const enabled = (id: CommandId): boolean => {
    switch (id) {
      case 'undo':
        return status.value.canUndo;
      case 'redo':
        return status.value.canRedo;
      // One of the pair at a time: the one that would change something.
      case 'iterate':
        return !status.value.iterating;
      case 'stopIterating':
        return status.value.iterating;
      case 'fillDown':
      case 'fillRight':
        // Always offered. Working out whether a fill would change
        // anything means reading the cells it would read, on the
        // wrong side of the barrier, to grey out a menu item.
        return true;
      default:
        return true;
    }
  };

  /** The active cell's paint, read without a round trip. */
  const paint = () => sheet.view.activeFormat.value.paint;

  /** The rows and columns the selection covers, corners normalised. */
  const rows = () => {
    const at = edit.selectionNow();
    const first = Math.min(cornerOf(at).row, at.anchorRow);
    const last = Math.max(cornerOf(at).row, at.anchorRow);
    return { first, last, count: last - first + 1 };
  };
  const columns = () => {
    const at = edit.selectionNow();
    const first = Math.min(cornerOf(at).column, at.anchorColumn);
    const last = Math.max(cornerOf(at).column, at.anchorColumn);
    return { first, last, count: last - first + 1 };
  };

  const format = (change: SheetFormatChange): void => sheet.send.format(change);

  /**
   * Sorting, with the block left to the other side to find.
   *
   * A selection of one cell means "sort the table I am standing in",
   * which is how everybody sorts. Where that table stops is not
   * something this thread can see — it holds the rows it has mounted
   * and nothing else — so the question crosses as a flag and the
   * application worker answers it.
   */
  const sortBy = (ascending: boolean): void => {
    const at = edit.selectionNow();
    const single = isOneCell(at);
    sheet.send.sortRange(single ? at.column : Math.min(cornerOf(at).column, at.anchorColumn), ascending, single);
  };

  /**
   * The theme in use, ticked when the View menu opens. Show formulas
   * and Show references say Hide while they are on, which is their
   * tick, so they have none.
   */
  const checkedNow = (id: CommandId): boolean | undefined => {
    switch (id) {
      case 'appearanceAuto':
        return appearance.value.value === 'auto';
      case 'appearanceLight':
        return appearance.value.value === 'light';
      case 'appearanceDark':
        return appearance.value.value === 'dark';
      default:
        return undefined;
    }
  };

  const run = (id: CommandId): void => {
    switch (id) {
      case 'undo':
        sheet.send.undo();
        break;
      case 'redo':
        sheet.send.redo();
        break;
      case 'cut':
        sheet.send.copy(true);
        break;
      case 'copy':
        sheet.send.copy(false);
        break;
      /**
       * The browser will not hand a worker the clipboard.
       *
       * `ShellService` can *write* it — that is what Phase 5's copy
       * goes through — and there is no matching read, because the
       * read is gated on a gesture in a document and the render
       * worker has no document. A menu item that silently did
       * nothing would be the worst of the three options, so this one
       * says what to press instead. Google Sheets does the same, for
       * the same reason.
       */
      case 'paste':
        pasteHint.value = true;
        break;
      case 'clear':
        sheet.send.clearRange();
        break;
      case 'pasteValues':
        sheet.send.pasteSpecial('values');
        break;
      case 'pasteFormats':
        sheet.send.pasteSpecial('formats');
        break;
      case 'pasteTransposed':
        sheet.send.pasteSpecial('transposed');
        break;
      // Lit for the next thing clicked; pressed again, put out.
      case 'formatPainter':
        if (edit.painter.value === 'off') {
          sheet.send.pickFormats();
          edit.setPainter('once');
        } else {
          edit.setPainter('off');
        }
        break;
      case 'selectAll':
        edit.apply({ kind: 'selectAll' });
        break;
      case 'fillDown':
        sheet.send.fillDown();
        break;
      case 'fillRight':
        sheet.send.fillRight();
        break;
      case 'find':
      case 'replace':
        askFor('find', () => (finding.value = id));
        return;
      case 'conditionalFormat':
      case 'dataValidation':
        askFor('rules', () => {
          rulesTab.value = id === 'conditionalFormat' ? 'format' : 'validation';
          ruling.value = true;
        });
        return;
      case 'insertChart':
        chartingForSelection = false;
        askFor('chart', () => (charting.value = true));
        return;
      case 'clearRules':
        sheet.send.clearRules();
        break;
      case 'gotoCell':
        notice.value = '';
        askFor('name');
        return;
      case 'newDocument':
        files.newDocument();
        break;
      case 'openFile':
        files.open();
        break;
      case 'openRecent':
        askFor('recent', () => (recenting.value = true));
        return;
      case 'saveDocument':
        files.save(false);
        break;
      case 'saveDocumentAs':
        files.save(true);
        break;
      case 'downloadCsv':
        files.exportCsv();
        break;
      case 'downloadXlsx':
        files.exportXlsx();
        break;
      /**
       * `Insert ▸ Name`, which is the name box with a sentence under
       * it.
       *
       * The item exists because the gesture had no way in: somebody
       * who has not been told that the box showing `B7` also defines
       * names will never find out. It does not open a dialog, because
       * a dialog would ask which range — and the selection has
       * already answered, which is the whole reason the gesture is
       * the box.
       */
      case 'defineName': {
        const at = edit.selectionNow();
        const range = formatRange({
          start: relativeRef(Math.min(cornerOf(at).row, at.anchorRow), Math.min(cornerOf(at).column, at.anchorColumn)),
          end: relativeRef(Math.max(cornerOf(at).row, at.anchorRow), Math.max(cornerOf(at).column, at.anchorColumn))
        });
        if (isOneCell(at)) {
          // Said rather than greyed out. A disabled item tells
          // somebody they cannot, and this tells them how.
          notice.value = ONE_CELL;
          return;
        }
        notice.value = `Type a name for ${range}, then press Enter.`;
        askFor('name');
        return;
      }
      /**
       * Guarded rather than trusted, because a command has three ways
       * in and only two of them are drawn. The button and the menu
       * item are both gone on `/`, and F9 is not — the key table is
       * one table for the whole application — so without this the
       * route would hide the command and answer it anyway.
       */
      // ---- the sheets --------------------------------------------
      case 'insertSheet':
        sheet.send.addSheet();
        break;
      case 'duplicateSheet':
        sheet.send.duplicateSheet(sheet.view.sheets.value.active);
        break;
      /**
       * Renaming happens in the strip along the bottom, which is a
       * sibling of this one — so the command goes through the handle
       * both of them hold rather than being done twice.
       */
      case 'renameSheet':
        edit.renameSheet();
        return;
      case 'sheetTabs':
        edit.focusTabs();
        return;
      case 'deleteSheet': {
        const tabs = sheet.view.sheets.value;
        if (tabs.entries.length > 1) {
          sheet.send.removeSheet(tabs.active);
        }
        break;
      }
      case 'moveSheetLeft':
      case 'moveSheetRight': {
        const tabs = sheet.view.sheets.value;
        const to = tabs.active + (id === 'moveSheetLeft' ? -1 : 1);
        if (to >= 0 && to < tabs.entries.length) {
          sheet.send.moveSheet(tabs.active, to);
        }
        break;
      }
      case 'nextSheet':
      case 'previousSheet': {
        const tabs = sheet.view.sheets.value;
        // Wrapped, because a strip somebody is stepping through has
        // two ends and stopping dead at one of them is a shortcut
        // that stops working exactly when it is being used.
        const count = tabs.entries.length;
        const to = (tabs.active + (id === 'nextSheet' ? 1 : -1) + count) % count;
        sheet.send.activateSheet(to);
        break;
      }
      case 'sheetColourNone':
      case 'sheetColourBlue':
      case 'sheetColourRed':
      case 'sheetColourGreen':
      case 'sheetColourPurple':
      case 'sheetColourOrange': {
        const named = id.slice('sheetColour'.length).toLowerCase();
        sheet.send.setSheetColour(sheet.view.sheets.value.active, TAB_COLOURS[named] ?? null);
        break;
      }
      case 'recalculate':
        if (proof) {
          sheet.send.stress(STRESS_CELLS);
        }
        break;
      case 'scriptStress':
        if (proof) {
          sheet.send.scriptStress();
        }
        break;
      case 'functionStress':
        if (proof) {
          sheet.send.functionStress(FUNCTION_CALLS);
        }
        break;
      case 'chartStress':
        if (proof) {
          sheet.send.chartStress(CHART_POINTS);
        }
        break;
      case 'shortcuts':
        shortcutsOpen.value = true;
        return;
      case 'scripts':
        scriptsOpen.value = true;
        return;
      case 'manageNames':
        namesOpen.value = true;
        return;
      case 'textColour':
        textPaletteOpen.value = true;
        return;
      case 'fillColour':
        fillPaletteOpen.value = true;
        return;
      case 'editNote': {
        const at = sheet.view.editor.value;
        noteAt = { row: at.row, column: at.column };
        noteCell.value = `${columnName(at.column)}${at.row + 1}`;
        noteText.value = at.note;
        noteOpen.value = true;
        return;
      }
      case 'menuBar':
        if (menuBarNode !== null) {
          focus.focus(menuBarNode);
        }
        return;
      /**
       * The three that toggle rather than set.
       *
       * Read off the *active cell*, because a toolbar has one Bold
       * button and a selection can hold both. Every spreadsheet
       * answers this from the active cell and this one does too: if
       * the cell you are on is bold, the button turns the selection
       * plain; if it is not, the button turns it bold.
       */
      case 'bold':
        format({ bold: !paint().bold });
        break;
      case 'italic':
        format({ italic: !paint().italic });
        break;
      case 'underline':
        format({ underline: !paint().underline });
        break;
      case 'wrap':
        format({ wrap: !paint().wrap });
        break;
      case 'alignLeft':
        format({ align: paint().align === 'start' ? 'auto' : 'start' });
        break;
      case 'alignCenter':
        format({ align: paint().align === 'center' ? 'auto' : 'center' });
        break;
      case 'alignRight':
        format({ align: paint().align === 'end' ? 'auto' : 'end' });
        break;
      case 'formatGeneral':
        format({ number: { kind: 'general' } });
        break;
      case 'formatNumber':
        format({ number: { kind: 'number', places: 2, thousands: true } });
        break;
      case 'formatCurrency':
        format({ number: { kind: 'currency', places: 2, symbol: '$' } });
        break;
      case 'formatPercent':
        format({ number: { kind: 'percent', places: 0 } });
        break;
      case 'formatScientific':
        format({ number: { kind: 'scientific', places: 2 } });
        break;
      case 'formatDate':
        format({ number: { kind: 'date', pattern: 'ymd' } });
        break;
      case 'formatTime':
        format({ number: { kind: 'time', pattern: 'hm' } });
        break;
      case 'formatDateTime':
        format({ number: { kind: 'datetime', date: 'ymd', time: 'hm' } });
        break;
      case 'formatText':
        format({ number: { kind: 'text' } });
        break;
      case 'moreDecimals':
        format({ places: 1 });
        break;
      case 'fewerDecimals':
        format({ places: -1 });
        break;
      case 'clearFormat':
        sheet.send.clearFormat();
        break;
      /**
       * How many rows the selection covers is how many go in.
       * Selecting three rows and asking for a row gives three, which
       * is what every spreadsheet does and saves the press-it-again
       * that people otherwise do anyway.
       */
      case 'insertRowAbove':
        sheet.send.insertRows(rows().first, rows().count);
        break;
      case 'insertRowBelow':
        sheet.send.insertRows(rows().last + 1, rows().count);
        break;
      case 'insertColumnLeft':
        sheet.send.insertColumns(columns().first, columns().count);
        break;
      case 'insertColumnRight':
        sheet.send.insertColumns(columns().last + 1, columns().count);
        break;
      case 'deleteRows':
        sheet.send.deleteRows(rows().first, rows().count);
        break;
      case 'deleteColumns':
        sheet.send.deleteColumns(columns().first, columns().count);
        break;
      case 'borderAll':
        sheet.send.setBorders('all', 1, '');
        break;
      case 'borderOutline':
        sheet.send.setBorders('outline', 1, '');
        break;
      case 'borderTop':
        sheet.send.setBorders('top', 1, '');
        break;
      case 'borderBottom':
        sheet.send.setBorders('bottom', 1, '');
        break;
      case 'borderThickBottom':
        sheet.send.setBorders('bottom', 2, '');
        break;
      case 'borderNone':
        sheet.send.setBorders('none', 0, '');
        break;
      case 'mergeCells':
        sheet.send.mergeCells();
        break;
      case 'unmergeCells':
        sheet.send.unmergeCells();
        break;
      case 'sortAscending':
        sortBy(true);
        break;
      case 'sortDescending':
        sortBy(false);
        break;
      case 'hideColumns':
        sheet.send.hideColumns(columns().first, columns().last);
        break;
      case 'showColumns':
        sheet.send.showColumns(columns().first, columns().last);
        break;
      case 'hideRows':
        sheet.send.hideRows(rows().first, rows().last);
        break;
      case 'showRows':
        sheet.send.showRows(rows().first, rows().last);
        break;
      case 'autofitColumns':
        sheet.send.measureColumns(columns().first, columns().last);
        break;
      case 'fitRows':
        sheet.send.fitRowsToContents(rows().first, rows().last);
        break;
      case 'showFormulas':
        sheet.send.showFormulas(!status.value.showingFormulas);
        break;
      case 'showReferences':
        edit.setReferencesShown(!edit.referencesShown.value);
        return;
      case 'zoomIn':
      case 'zoomOut':
      case 'zoomReset':
        sheet.send.setZoom(zoomStep(sheet.view.geometry.value.zoom, id));
        break;
      case 'appearanceAuto':
        appearance.set('auto');
        break;
      case 'appearanceLight':
        appearance.set('light');
        break;
      case 'appearanceDark':
        appearance.set('dark');
        break;
      case 'iterate':
        sheet.send.setIteration(true);
        break;
      case 'stopIterating':
        sheet.send.setIteration(false);
        break;
      case 'filterToSelection':
        sheet.send.filterToSelection();
        break;
      case 'clearFilter':
        sheet.send.clearFilter();
        break;
      /**
       * "Up to here" means everything above and to the left of the
       * active cell, which is what the phrase means to a person and
       * saves them counting rows.
       */
      case 'freezeHere':
        sheet.send.freeze(edit.selectionNow().row, edit.selectionNow().column);
        break;
      case 'freezeTopRow':
        sheet.send.freeze(1, 0);
        break;
      case 'freezeFirstColumn':
        sheet.send.freeze(0, 1);
        break;
      /**
       * The right-click on a letter or a number: freeze through the
       * rows or columns selected, including them, and leave the other
       * direction's freeze as it was.
       */
      case 'freezeThroughRows':
        sheet.send.freeze(rows().last + 1, sheet.view.geometry.value.frozenColumns);
        break;
      case 'freezeThroughColumns':
        sheet.send.freeze(sheet.view.geometry.value.frozenRows, columns().last + 1);
        break;
      case 'unfreeze':
        sheet.send.freeze(0, 0);
        break;
      case 'unfreezeRows':
        sheet.send.freeze(0, sheet.view.geometry.value.frozenColumns);
        break;
      case 'unfreezeColumns':
        sheet.send.freeze(sheet.view.geometry.value.frozenRows, 0);
        break;
    }
    edit.focusSheet();
  };

  // The grid holds focus while somebody is using the sheet, so the
  // accelerators arrive there; this is what they reach.
  edit.provideCommands(run);

  /**
   * What Escape closes, innermost first.
   *
   * A dialog before the find bar, because a dialog is over the top of
   * it: closing the bar underneath a dialog would be answering a key
   * with the thing the person cannot see.
   */
  edit.provideDismiss(() => {
    if (noteOpen.value) {
      closeNote();
      return true;
    }
    if (scriptsOpen.value) {
      closeScripts();
      return true;
    }
    if (namesOpen.value) {
      closeNames();
      return true;
    }
    if (customFor.value !== null) {
      closeCustom();
      return true;
    }
    if (shortcutsOpen.value) {
      shortcutsOpen.value = false;
      return true;
    }
    if (pasteHint.value) {
      pasteHint.value = false;
      return true;
    }
    if (recenting.value) {
      recenting.value = false;
      return true;
    }
    if (charting.value) {
      // A chart keeps its handles when the bar closes; Escape here is
      // about the bar, and clicking the grid is what lets a chart go.
      charting.value = false;
      return true;
    }
    if (ruling.value) {
      ruling.value = false;
      return true;
    }
    if (finding.value !== 'closed') {
      sheet.send.clearFind();
      finding.value = 'closed';
      return true;
    }
    return false;
  });

  /**
   * The find bar, built once and switched in and out.
   *
   * Built inside the `map` it would be a *new* element on every
   * emission, and a new element is a new component, a new node and a
   * new field — so Ctrl+F focused one instance and the person typed
   * into the one that replaced it. Elements are values here; handing
   * back the same one is what makes switching it in cheap and makes
   * the thing inside it keep its state. The grid memoises its cells
   * for the same reason, and it is the same rule.
   */
  const findBar = [
    <box key="rule" width={percent(100)} height={1} backgroundColor="border" />,
    <FindBar
      key="find"
      editing={edit}
      replacing={finding.pipe(map(current => current === 'replace'))}
      onClose={() => (finding.value = 'closed')}
      ref={arrived('find')}
    />
  ];

  /**
   * The notice, as a rule and a line of text.
   *
   * Rebuilt per sentence rather than built once like the find bar,
   * which is safe here for the reason it was not there: there is
   * nothing inside it to focus and nothing to lose. It carries
   * `status` so a screen reader is told without being interrupted,
   * which is what a sentence under a field is for.
   */
  const noticeRow = (text: string) => [
    <box key="rule" width={percent(100)} height={1} backgroundColor="border" />,
    <row
      key="notice"
      width={percent(100)}
      y="center"
      paddingLeft={8}
      paddingRight={8}
      paddingTop={4}
      paddingBottom={4}
      backgroundColor="surface"
      role="status"
      label="Name box notice">
      <text text={text} fontSize={11} color="textMuted" textWrap="word" selectable={false} />
    </row>
  ];

  /**
   * The rules bar, built once and switched in — for the reason the
   * find bar is, and it was written here the other way first.
   *
   * Built inside the `map` it is a *new* element on every emission,
   * which is a new component, a new node and a new set of fields.
   * The bar mounted, its roles reached the accessibility tree, and
   * every node in it had a width and a height of zero: it was being
   * replaced before it had ever been laid out. Every spec passed,
   * because a spec asks what a node's properties are and the
   * properties were right.
   */
  /**
   * The bar that says a file's functions are off — Phase 33.
   *
   * Shown while the workbook has function scripts from a file that
   * nobody has turned on, and until somebody says Not now. Built once
   * and switched in, for the reason the rules bar below is. Turn on is
   * here and in the editor; the bar is where somebody who has only
   * opened a file and seen `#NAME?` will look.
   */
  const functionsFrom = sheet.view.scripts.pipe(
    map(scripts => scripts.entries.find(entry => entry.kind === 'functions' && entry.from !== '' && !entry.on)?.from ?? ''),
    distinctUntilChanged()
  );
  const functionsDismissed = internalState('');
  const functionsShown = combineLatest([functionsFrom, functionsDismissed]).pipe(
    map(([from, dismissed]) => from !== '' && from !== dismissed),
    distinctUntilChanged()
  );
  const barButton = (label: string, onClick: () => void) => (
    <button
      onClick={onClick}
      label={label}
      paddingLeft={9}
      paddingRight={9}
      paddingTop={3}
      paddingBottom={3}
      borderRadius={5}
      backgroundColor="controlBackground"
      borderColor="controlBorder"
      borderWidth={1}
      cursor="pointer">
      <text text={label} fontSize={11} textWrap="none" color="controlForeground" selectable={false} />
    </button>
  );
  const functionsBar = [
    <box key="rule" width={percent(100)} height={1} backgroundColor="border" />,
    <row
      key="functions"
      width={percent(100)}
      y="center"
      gap={8}
      paddingLeft={8}
      paddingRight={8}
      paddingTop={4}
      paddingBottom={4}
      backgroundColor="surface"
      role="region"
      label="Functions from a file">
      <text
        text={functionsFrom.pipe(
          map(from => `This workbook's functions came with ${from}, and are off, so a formula that calls one says #NAME?. Turn them on only for a file you trust.`)
        )}
        flex={1}
        minWidth={0}
        fontSize={11}
        color="text"
        textWrap="word"
        selectable={false}
      />
      {barButton('Turn on', () => sheet.send.setFunctionsOn(true))}
      {barButton('Scripts…', () => (scriptsOpen.value = true))}
      {barButton('Not now', () => (functionsDismissed.value = sheet.view.scripts.value.entries.find(entry => entry.kind === 'functions' && entry.from !== '' && !entry.on)?.from ?? ''))}
    </row>
  ];

  const rulesBar = [
    <box key="rule" width={percent(100)} height={1} backgroundColor="border" />,
    <RulesBar
      key="rules"
      editing={edit}
      tab={rulesTab}
      onTab={(next: RulesTab) => (rulesTab.value = next)}
      onClose={() => (ruling.value = false)}
      ref={arrived('rules')}
    />
  ];

  /**
   * The chart bar, built once and switched in — for the reason the
   * find bar and the rules bar are.
   */
  /** The recent-files bar, built once and switched in, for the reason the others are. */
  const recentBar = [
    <box key="recentrule" width={percent(100)} height={1} backgroundColor="border" />,
    <RecentBar
      key="recent"
      files={files}
      onClose={() => {
        recenting.value = false;
        edit.focusSheet();
      }}
      ref={arrived('recent')}
    />
  ];

  const chartBar = [
    <box key="chartrule" width={percent(100)} height={1} backgroundColor="border" />,
    <ChartBar key="chart" editing={edit} onClose={() => (charting.value = false)} ref={arrived('chart')} />
  ];

  /**
   * What the formula bar shows: the draft while a cell is open, and
   * what the application worker says the cell holds otherwise.
   *
   * The same buffer as the cell, not a copy of it. Two buffers kept
   * in step would be two answers to what Escape puts back, and the
   * bar and the cell would disagree for exactly as long as it took a
   * keystroke to cross between them.
   */
  const formula: Observable<string> = combineLatest([edit.draft, sheet.view.editor]).pipe(
    map(([draft, current]) => draft ?? current.spilledFrom?.input ?? current.input)
  );

  /**
   * Whether the bar is showing a formula that is not this cell's: the
   * cursor is on a cell an array spilled into. Greyed, and read-only —
   * an edit begun in the bar would start from the formula and write it
   * into the spilled cell, which blocks the array it came from. Typing
   * over the cell in the grid still replaces it, as in Excel.
   */
  const borrowed: Observable<boolean> = combineLatest([edit.draft, sheet.view.editor]).pipe(
    map(([draft, current]) => draft === null && current.spilledFrom !== null),
    distinctUntilChanged()
  );

  /**
   * The formula bar's references, in the colours the grid draws their
   * boxes in.
   *
   * The same `formulaSpans` the cell editor uses, so the `B2` in the
   * bar, the `B2` in the cell and the box round B2 on the sheet are
   * one colour by construction rather than by three files agreeing.
   *
   * Coloured whether or not the bar has the caret, because a selected
   * cell showing `=SUM(Sales)+B2` is worth reading at a glance and
   * the colours are most of what makes it readable. The *bracket*
   * marking needs a caret and gets one only while the bar is being
   * typed in, which is the only time it means anything.
   */
  const formulaSpans$ = new BehaviorSubject<readonly UiTextSpan[] | undefined>(undefined);
  let formulaNode: UiNode | null = null;
  let formulaText = '';

  const refreshFormulaSpans = (): void => {
    const caret =
      formulaNode !== null && focus.focused.value === formulaNode ? editorFor(formulaNode).focus : undefined;
    formulaSpans$.next(formulaSpans(formulaText, caret));
  };

  ctx.effect(formula, text => {
    formulaText = text;
    refreshFormulaSpans();
  });

  const onFormulaKey = (event: UiKeyboardEvent): void => {
    // Always read as "a cell is open", whatever the draft says,
    // because the caret is in a text field and the keys belong to the
    // text. Read the other way — as "a cell is selected and nothing
    // is open" — Backspace meant *empty this cell* and a digit meant
    // *replace this cell*, so deleting one character wiped the lot.
    if (edit.apply(keyAction(event.key, event.modifiers, true))) {
      event.preventDefault();
      event.stopPropagation();
    }
  };

  return (
    <column width={percent(100)} flexShrink={0} backgroundColor="surface">
      <row width={percent(100)} y="center" paddingTop={3} paddingBottom={3}>
        <MenuBar
          menus={menusFor(proof)}
          enabled={enabled}
          labelNow={labelNow}
          checkedNow={checkedNow}
          onChoose={run}
          /**
           * The sheet gets the keyboard back, unless the command that
           * closed the menu wanted it somewhere else.
           *
           * `Insert ▸ Name` puts it in the name box and the menu
           * closes immediately after — so without this the box was
           * focused and then unfocused within the frame, and the
           * person typed their name into cell A1. `wanted` covers the
           * field that has not mounted yet, which is the find bar's
           * case.
           */
          onDismiss={() => {
            const claimed = focus.focused.value;
            if (
              wanted === null &&
              claimed !== nameBoxNode &&
              claimed !== findFieldNode &&
              claimed !== rulesFieldNode &&
              claimed !== chartFieldNode &&
              claimed !== recentFieldNode
            ) {
              edit.focusSheet();
            }
          }}
          ref={(node: UiNode | null) => (menuBarNode = node)}
        />
        <box flex={1} minWidth={0} />
        <Toolbar items={tools} label="Formatting" />
        <box width={8} />
      </row>
      <box width={percent(100)} height={1} backgroundColor="border" />
      <row width={percent(100)} y="center" gap={8} padding={6}>
        <NameBox editing={edit} ref={arrived('name')} onNotice={(text: string) => (notice.value = text)} />
        <editabletext
          value={formula}
          spans={formulaSpans$}
          ref={(node: UiNode | null) => {
            formulaNode = node;
            refreshFormulaSpans();
          }}
          flex={1}
          minWidth={0}
          fontSize={12}
          color={borrowed.pipe(map(is => (is ? 'textMuted' : 'text')))}
          readOnly={borrowed}
          textWrap="none"
          verticalAlign="middle"
          backgroundColor="background"
          borderColor="border"
          borderWidth={1}
          padding={4}
          role="textbox"
          label="Formula"
          onInput={(event: UiTextChangeEvent) => edit.write(event.value)}
          onKeyDown={onFormulaKey}
          // The caret moving with the text unchanged, which is what
          // the bracket beside it depends on.
          onSelectionChange={refreshFormulaSpans}
          onFocus={refreshFormulaSpans}
          onBlur={refreshFormulaSpans}
        />
      </row>
      {ruling.pipe(map(open => (open ? rulesBar : [])))}
      {charting.pipe(map(open => (open ? chartBar : [])))}
      {recenting.pipe(map(open => (open ? recentBar : [])))}
      {notice.pipe(map(text => (text === '' ? [] : noticeRow(text))))}
      {functionsShown.pipe(map(open => (open ? functionsBar : [])))}
      {finding.pipe(map(mode => (mode === 'closed' ? [] : findBar)))}
      <Shortcuts proof={proof} open={shortcutsOpen} onClose={() => (shortcutsOpen.value = false)} />
      <PasteHint open={pasteHint} onClose={() => (pasteHint.value = false)} />
      <NoteDialog
        open={noteOpen}
        cell={noteCell}
        text={noteText}
        onSave={(text: string) => {
          sheet.send.setNote(noteAt.row, noteAt.column, text);
          closeNote();
        }}
        onClose={closeNote}
      />
      <ScriptDialog
        open={scriptsOpen}
        scripts={sheet.view.scripts}
        onSave={(was: string, name: string, source: string, kind: 'run' | 'functions') => sheet.send.saveScript(was, name, source, kind)}
        onRemove={(name: string) => sheet.send.removeScript(name)}
        onRun={(name: string, confirmed: boolean) => sheet.send.runScript(name, confirmed)}
        onStop={() => sheet.send.stopScript()}
        onFunctionsOn={(on: boolean) => sheet.send.setFunctionsOn(on)}
        onClose={closeScripts}
      />
      <ColorPalette
        open={textPaletteOpen}
        onOpenChange={(open: boolean) => (textPaletteOpen.value = open)}
        anchor={textAnchor}
        value={sheet.view.activeFormat.pipe(map(current => current.paint.color))}
        recent={recentColours}
        automaticLabel="Automatic"
        label="Text colour"
        onSelect={(colour: string) => paintWith({ color: colour })}
        onCustom={() => openCustom('color')}
      />
      <ColorPalette
        open={fillPaletteOpen}
        onOpenChange={(open: boolean) => (fillPaletteOpen.value = open)}
        anchor={fillAnchor}
        value={sheet.view.activeFormat.pipe(map(current => current.paint.fill))}
        recent={recentColours}
        automaticLabel="No fill"
        label="Fill colour"
        onSelect={(colour: string) => paintWith({ fill: colour })}
        onCustom={() => openCustom('fill')}
      />
      <Dialog
        open={customFor.pipe(map(which => which !== null))}
        onClose={closeCustom}
        title={customFor.pipe(map(which => (which === 'fill' ? 'Custom fill colour' : 'Custom text colour')))}
        width={260}
        content={
          <column gap={12}>
            <ColorPicker
              value={customColour}
              label={customFor.pipe(map(which => (which === 'fill' ? 'Fill colour' : 'Text colour')))}
              width={224}
              onChange={(colour: string) => (customColour.value = colour)}
            />
            <row gap={8} x="end">
              <button
                onClick={closeCustom}
                label="Cancel"
                paddingLeft={12}
                paddingRight={12}
                paddingTop={5}
                paddingBottom={5}
                borderRadius={6}
                backgroundColor="controlBackground"
                borderColor="controlBorder"
                borderWidth={1}
                cursor="pointer">
                <text text="Cancel" fontSize={12} color="controlForeground" selectable={false} />
              </button>
              <button
                onClick={() => {
                  const which = customFor.value;
                  customFor.value = null;
                  if (which !== null) {
                    paintWith(which === 'fill' ? { fill: customColour.value } : { color: customColour.value });
                  }
                }}
                label="Use colour"
                paddingLeft={12}
                paddingRight={12}
                paddingTop={5}
                paddingBottom={5}
                borderRadius={6}
                backgroundColor="primary"
                borderColor="controlBorder"
                borderWidth={1}
                cursor="pointer">
                <text text="Use colour" fontSize={12} color="primaryForeground" selectable={false} />
              </button>
            </row>
          </column>
        }
      />
      <NamesDialog
        open={namesOpen}
        names={sheet.view.names}
        onSave={(was: string, name: string, refersTo: string) => sheet.send.saveName(was, name, refersTo)}
        onRemove={(name: string) => sheet.send.removeName(name)}
        onClose={closeNames}
      />
    </column>
  );
}

