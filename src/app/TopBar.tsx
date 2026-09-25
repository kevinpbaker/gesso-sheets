import { combineLatest, map, type Observable } from 'rxjs';

import { editorFor, percent, type UiKeyboardEvent, type UiNode, type UiTextChangeEvent } from 'gesso-core';
import { FocusService, internalState, type ComponentContext, type Inputs } from 'gesso-framework';

import { FindBar } from './FindBar';
import { MenuBar } from './MenuBar';
import { NameBox } from './NameBox';
import { Sheet } from './SheetContract';
import { COMMANDS, STRESS_CELLS, type CommandId } from './SheetCommands';
import type { SheetEditing } from './SheetEditing';
import { keyAction } from './SheetKeys';
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
}

/** Which half of the find bar is showing, or neither. */
type Finding = 'closed' | 'find' | 'replace';

export function TopBar(inputs: Inputs<TopBarProps>, ctx: ComponentContext) {
  const sheet = ctx.channel(Sheet);
  const focus = ctx.inject(FocusService);
  const edit = inputs.editing.value;
  const status = sheet.view.status;

  const finding = internalState<Finding>('closed');
  const shortcutsOpen = internalState(false);
  /** Where the paste hint is, when somebody asks for Paste from a menu. */
  const pasteHint = internalState(false);

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
  let wanted: 'name' | 'find' | null = null;

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
  const askFor = (which: 'name' | 'find', open?: () => void): void => {
    wanted = which;
    open?.();
    const node = which === 'name' ? nameBoxNode : findFieldNode;
    if (wanted === which && node !== null) {
      wanted = null;
      take(node);
    }
  };

  const arrived = (which: 'name' | 'find') => (node: UiNode | null) => {
    if (which === 'name') {
      nameBoxNode = node;
    } else {
      findFieldNode = node;
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
      case 'gotoCell':
        askFor('name');
        return;
      case 'recalculate':
        sheet.send.stress(STRESS_CELLS);
        break;
      case 'shortcuts':
        shortcutsOpen.value = true;
        return;
      case 'menuBar':
        if (menuBarNode !== null) {
          focus.focus(menuBarNode);
        }
        return;
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
    if (shortcutsOpen.value) {
      shortcutsOpen.value = false;
      return true;
    }
    if (pasteHint.value) {
      pasteHint.value = false;
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
   * What the formula bar shows: the draft while a cell is open, and
   * what the application worker says the cell holds otherwise.
   *
   * The same buffer as the cell, not a copy of it. Two buffers kept
   * in step would be two answers to what Escape puts back, and the
   * bar and the cell would disagree for exactly as long as it took a
   * keystroke to cross between them.
   */
  const formula: Observable<string> = combineLatest([edit.draft, sheet.view.editor]).pipe(
    map(([draft, current]) => draft ?? current.input)
  );

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
          enabled={enabled}
          onChoose={run}
          onDismiss={() => edit.focusSheet()}
          ref={(node: UiNode | null) => (menuBarNode = node)}
        />
        <box flex={1} minWidth={0} />
        {toolButton('Undo', status.pipe(map(s => s.canUndo)), () => run('undo'))}
        {toolButton('Redo', status.pipe(map(s => s.canRedo)), () => run('redo'))}
        {toolButton(COMMANDS.recalculate.label, undefined, () => run('recalculate'))}
        <box width={8} />
      </row>
      <box width={percent(100)} height={1} backgroundColor="border" />
      <row width={percent(100)} y="center" gap={8} padding={6}>
        <NameBox editing={edit} ref={arrived('name')} />
        <editabletext
          value={formula}
          flex={1}
          minWidth={0}
          fontSize={12}
          color="text"
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
        />
      </row>
      {finding.pipe(map(mode => (mode === 'closed' ? [] : findBar)))}
      <Shortcuts open={shortcutsOpen} onClose={() => (shortcutsOpen.value = false)} />
      <PasteHint open={pasteHint} onClose={() => (pasteHint.value = false)} />
    </column>
  );
}

function toolButton(label: string, enabled: Observable<boolean> | undefined, onClick: () => void) {
  return (
    <button
      onClick={onClick}
      label={label}
      paddingLeft={10}
      paddingRight={10}
      paddingTop={4}
      paddingBottom={4}
      marginLeft={4}
      borderRadius={6}
      backgroundColor="controlBackground"
      borderColor="controlBorder"
      borderWidth={1}
      cursor="pointer"
      opacity={enabled === undefined ? 1 : enabled.pipe(map(on => (on ? 1 : 0.4)))}>
      <text text={label} fontSize={12} color="controlForeground" selectable={false} />
    </button>
  );
}
