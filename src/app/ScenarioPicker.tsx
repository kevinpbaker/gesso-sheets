import { BehaviorSubject, combineLatest, map } from 'rxjs';

import { Column, Text, type UiKeyboardEvent, type UiNode, type UiTextChangeEvent } from 'gesso-core';
import { Dialog, Menu, type MenuItem } from 'gesso-components';
import { createComponent, FocusService, internalState, type ComponentContext, type Inputs } from 'gesso-framework';

import { Sheet, type SheetScenarios } from './SheetContract';
import type { SheetEditing } from './SheetEditing';
import { SCENARIO_CHANGED, SCENARIO_TYPED } from './scenarioPaint';

/**
 * Which scenario the sheet is showing, beside the tabs.
 *
 * Beside the tabs because it is the same kind of choice: which version
 * of the workbook you are looking at. A button that says which, and a
 * menu that lists Base and every scenario, with the three things done
 * to one — make one, rename it, delete it — and the one thing done to
 * cells in it: give them back to the base.
 *
 * While a scenario is showing, the strip also says how the active cell
 * stands against the base. That is the one question the tint cannot
 * answer — it says a cell moved, and not from what.
 */
export interface ScenarioPickerProps {
  readonly editing: SheetEditing;
}

/** The menu's values that are not a scenario's id. */
const BASE = '\u0000base';
const NEW = '\u0000new';
const COPY = '\u0000copy';
const RENAME = '\u0000rename';
const DELETE = '\u0000delete';
const RESET = '\u0000reset';

export function ScenarioPicker(inputs: Inputs<ScenarioPickerProps>, ctx: ComponentContext) {
  const sheet = ctx.channel(Sheet);
  const focus = ctx.inject(FocusService);
  const edit = inputs.editing.value;
  const scenarios = sheet.view.scenarios;

  const shownName = (view: SheetScenarios): string => view.entries.find(entry => entry.id === view.shown)?.name ?? 'Base';

  const menuOpen = new BehaviorSubject(false);
  const menuAt = new BehaviorSubject({ x: 0, y: 0 });
  const items = scenarios.pipe(
    map((view): readonly MenuItem[] => {
      const mark = (on: boolean): string => (on ? '✓ ' : '   ');
      const showing = view.shown !== null;
      return [
        { value: BASE, label: `${mark(!showing)}Base` },
        ...view.entries.map(entry => ({
          value: entry.id,
          label: `${mark(entry.id === view.shown)}${entry.name}${entry.inputs === 0 ? '' : ` (${entry.inputs} ${entry.inputs === 1 ? 'input' : 'inputs'})`}`
        })),
        { value: NEW, label: 'New scenario…' },
        { value: COPY, label: 'New scenario from this one…', disabled: !showing },
        { value: RENAME, label: 'Rename this scenario…', disabled: !showing },
        { value: DELETE, label: 'Delete this scenario…', disabled: !showing },
        { value: RESET, label: 'Use the base for the selected cells', disabled: !showing }
      ];
    })
  );

  /** The name dialog: what it is for, and what has been typed so far. */
  const naming = internalState<'new' | 'copy' | 'rename' | null>(null);
  const draft = internalState('');
  const deleting = internalState(false);

  const now = (): SheetScenarios => {
    let view: SheetScenarios = { entries: [], shown: null };
    scenarios.subscribe(next => (view = next)).unsubscribe();
    return view;
  };

  const choose = (value: string): void => {
    menuOpen.next(false);
    const view = now();
    switch (value) {
      case BASE:
        sheet.send.showScenario(null);
        break;
      case NEW:
        draft.value = view.entries.length === 0 ? 'Optimistic' : `Scenario ${view.entries.length + 1}`;
        naming.value = 'new';
        return;
      case COPY:
        draft.value = `${shownName(view)} copy`;
        naming.value = 'copy';
        return;
      case RENAME:
        draft.value = shownName(view);
        naming.value = 'rename';
        return;
      case DELETE:
        deleting.value = true;
        return;
      case RESET:
        sheet.send.resetScenarioCells();
        break;
      default:
        sheet.send.showScenario(value);
    }
    edit.focusSheet();
  };

  const menu = createComponent(Menu, {
    open: menuOpen,
    onOpenChange: (open: boolean) => menuOpen.next(open),
    items,
    at: menuAt,
    label: 'Scenarios',
    onSelect: choose
  });

  const button = ctx.bounds('scenario');
  const openMenu = (): void => {
    const box = button.value;
    menuAt.next({ x: box.x, y: box.y });
    menuOpen.next(true);
  };

  const closeNaming = (): void => {
    naming.value = null;
    edit.focusSheet();
  };
  const saveName = (): void => {
    const name = draft.value.trim();
    const what = naming.value;
    const view = now();
    if (name === '' || what === null) {
      return;
    }
    if (what === 'rename') {
      if (view.shown !== null) {
        sheet.send.renameScenario(view.shown, name);
      }
    } else {
      sheet.send.addScenario(name, what === 'copy');
    }
    closeNaming();
  };
  const closeDelete = (): void => {
    deleting.value = false;
    edit.focusSheet();
  };
  const confirmDelete = (): void => {
    const view = now();
    if (view.shown !== null) {
      sheet.send.deleteScenario(view.shown);
    }
    closeDelete();
  };

  /**
   * How the active cell stands against the base: what it was there, and
   * whether the scenario typed it or only moved it. The swatch is the
   * tint the cell is drawn in, so the line reads as the key to it.
   */
  const against = sheet.view.editor.pipe(
    map(editor => editor.scenario),
    map(cell =>
      cell === null
        ? []
        : [
            <row key="against" gap={5} y="center" minWidth={0}>
              <box width={10} height={10} borderRadius={2} backgroundColor={cell.typed ? SCENARIO_TYPED : SCENARIO_CHANGED} />
              <text
                text={`${cell.typed ? 'Typed in' : 'Changed by'} ${cell.name} · Base: ${cell.base === '' ? 'empty' : cell.base}`}
                fontSize={11}
                color="textMuted"
                verticalAlign="middle"
                selectable={false}
                textWrap="none"
                live="polite"
              />
            </row>
          ]
    )
  );

  const label = scenarios.pipe(map(view => `Scenario: ${shownName(view)}`));
  const showing = scenarios.pipe(map(view => view.shown !== null));

  const dialogButton = (text: string, onClick: () => void) => (
    <button
      onClick={onClick}
      label={text}
      paddingLeft={12}
      paddingRight={12}
      paddingTop={5}
      paddingBottom={5}
      borderRadius={6}
      backgroundColor="controlBackground"
      borderColor="controlBorder"
      borderWidth={1}
      cursor="pointer">
      <text text={text} fontSize={12} color="controlForeground" selectable={false} />
    </button>
  );

  const onNameKey = (event: UiKeyboardEvent): void => {
    if (event.key === 'Enter') {
      event.preventDefault();
      event.stopPropagation();
      saveName();
    }
  };

  return (
    <row gap={10} y="center" minWidth={0}>
      {against}
      <button
        modifiers={[button.modifier]}
        onClick={openMenu}
        label={label}
        paddingLeft={9}
        paddingRight={9}
        paddingTop={3}
        paddingBottom={3}
        borderRadius={4}
        backgroundColor={showing.pipe(map(on => (on ? SCENARIO_TYPED : 'surface')))}
        borderColor={showing.pipe(map(on => (on ? SCENARIO_TYPED : 'border')))}
        borderWidth={1}
        cursor="pointer">
        <text
          text={label.pipe(map(text => `${text} ▾`))}
          fontSize={11}
          color={showing.pipe(map(on => (on ? '#1f2937' : 'text')))}
          selectable={false}
          textWrap="none"
        />
      </button>
      {menu}
      <Dialog
        open={naming.pipe(map(what => what !== null))}
        onClose={closeNaming}
        title={naming.pipe(map(what => (what === 'rename' ? 'Rename scenario' : 'New scenario')))}
        width={340}
        content={Column(
          { gap: 12, minWidth: 0 },
          Text({
            text: naming.pipe(
              map(what =>
                what === 'rename'
                  ? 'What to call it.'
                  : what === 'copy'
                    ? 'It starts with every input this one types, and shows straight away.'
                    : 'It starts as the base and shows straight away. What you type while it shows is its own, and the base keeps its values.'
              )
            ),
            fontSize: 12,
            color: 'textMuted',
            textWrap: 'word'
          }),
          <editabletext
            value={draft}
            ref={(node: UiNode | null) => {
              if (node !== null) {
                focus.focus(node);
              }
            }}
            fontSize={12}
            color="text"
            backgroundColor="background"
            borderColor="border"
            borderWidth={1}
            padding={6}
            role="textbox"
            label="Scenario name"
            onInput={(event: UiTextChangeEvent) => (draft.value = event.value)}
            onKeyDown={onNameKey}
          />,
          <row gap={8} x="end">
            {dialogButton('Cancel', closeNaming)}
            {dialogButton('Save', saveName)}
          </row>
        )}
      />
      <Dialog
        open={deleting}
        onClose={closeDelete}
        title="Delete scenario"
        width={340}
        content={Column(
          { gap: 12, minWidth: 0 },
          Text({
            text: combineLatest([scenarios, deleting]).pipe(
              map(([view]) => `${shownName(view)} and every input it types will go. The base is not changed.`)
            ),
            fontSize: 12,
            color: 'text',
            textWrap: 'word'
          }),
          <row gap={8} x="end">
            {dialogButton('Cancel', closeDelete)}
            {dialogButton('Delete', confirmDelete)}
          </row>
        )}
      />
    </row>
  );
}
