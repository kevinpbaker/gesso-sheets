import { map, type Observable } from 'rxjs';

import { percent, type UiKeyboardEvent, type UiNode, type UiSemanticState, type UiTextChangeEvent } from 'gesso-core';
import { internalState, type ComponentContext, type Inputs } from 'gesso-framework';

import type { ChartKind } from '../sheet/Chart';
import { Sheet } from './SheetContract';
import type { SheetEditing } from './SheetEditing';

/**
 * The bar that makes a chart, and then edits the one it made.
 *
 * A row in the flow rather than a floating panel, for the reason the
 * find bar and the rules bar are: opening it must not cover the range
 * the chart is of. It costs a row of height, which is the cheaper of
 * the two mistakes.
 *
 * **One bar for both jobs**, and that is worth saying because the
 * alternative was two. With nothing selected it inserts; with a chart
 * selected the same controls edit that chart, live. They are the same
 * question — what kind, called what, with a legend or not — and a
 * separate "chart properties" panel would be a second place to learn
 * the same four controls.
 *
 * Selecting a chart is a click on the chart, which means the bar has
 * to follow a selection it does not own. It reads `charts.selected`
 * off the channel and re-reads its own controls from whatever that
 * chart holds, so clicking between two charts moves the bar to the
 * second one rather than applying the first one's settings to it.
 */

export interface ChartBarProps {
  readonly editing: SheetEditing;
  readonly onClose: () => void;
  readonly ref?: (node: UiNode | null) => void;
}

/**
 * The seven kinds, in the order somebody looks for them.
 *
 * Column first because it is what most people mean by "chart", and
 * the pair that is one chart on its side — column and bar — next to
 * each other so the difference is visible rather than remembered.
 */
const KINDS: readonly { readonly id: ChartKind; readonly label: string }[] = [
  { id: 'column', label: 'Column' },
  { id: 'bar', label: 'Bar' },
  { id: 'stacked', label: 'Stacked' },
  { id: 'line', label: 'Line' },
  { id: 'area', label: 'Area' },
  { id: 'pie', label: 'Pie' },
  { id: 'scatter', label: 'Scatter' }
];

export function ChartBar(inputs: Inputs<ChartBarProps>, ctx: ComponentContext) {
  const sheet = ctx.channel(Sheet);
  const edit = inputs.editing.value;

  const kind = internalState<ChartKind>('column');
  const title = internalState('');
  /** The chart the bar is editing, or zero while it is an insert bar. */
  const editingId = internalState(0);

  /**
   * The bar follows the chart somebody clicked.
   *
   * Only when the *id* changes, which is the whole of the guard: the
   * charts key also moves on every drag, and reading the title back
   * out of it on each one would fight a person typing a new title
   * into the field while the chart was still moving.
   */
  ctx.effect(sheet.view.charts, view => {
    if (view.selected === editingId.value) {
      return;
    }
    editingId.value = view.selected;
    const chosen = view.entries.find(entry => entry.id === view.selected);
    if (chosen !== undefined) {
      kind.value = chosen.kind;
      title.value = chosen.title;
    }
  });

  const close = (): void => {
    inputs.onClose.value();
    edit.focusSheet();
  };

  const pickKind = (next: ChartKind): void => {
    kind.value = next;
    if (editingId.value !== 0) {
      sheet.send.setChartKind(editingId.value, next);
    }
  };

  const applyTitle = (): void => {
    if (editingId.value !== 0) {
      sheet.send.setChartTitle(editingId.value, title.value);
    }
  };

  const insert = (): void => {
    sheet.send.insertChart(kind.value);
    // The title in the field belongs to the chart being made, so it
    // is applied to whatever came back — which the effect above has
    // already told us the id of, because inserting selects.
    if (title.value !== '') {
      applyTitle();
    }
  };

  const remove = (): void => {
    if (editingId.value !== 0) {
      sheet.send.removeChart(editingId.value);
    }
  };

  const onFieldKey = (run: () => void) => (event: UiKeyboardEvent) => {
    if (event.key === 'Enter') {
      run();
      event.preventDefault();
      event.stopPropagation();
      return;
    }
    if (event.key === 'Escape') {
      close();
      event.preventDefault();
      event.stopPropagation();
    }
  };

  const button = (label: string, onClick: () => void, enabled: Observable<boolean> | true = true) => (
    <button
      key={`button-${label}`}
      onClick={onClick}
      label={label}
      disabled={enabled === true ? false : enabled.pipe(map(is => !is))}
      paddingLeft={9}
      paddingRight={9}
      paddingTop={4}
      paddingBottom={4}
      borderRadius={5}
      backgroundColor="controlBackground"
      borderColor="controlBorder"
      borderWidth={1}
      cursor="pointer">
      <text text={label} fontSize={11} textWrap="none" color="controlForeground" selectable={false} />
    </button>
  );

  const hasChart = editingId.pipe(map(id => id !== 0));

  return (
    <row
      width={percent(100)}
      flexShrink={0}
      y="center"
      gap={8}
      paddingLeft={10}
      paddingRight={10}
      paddingTop={5}
      paddingBottom={5}
      backgroundColor="surface"
      role="form"
      label="Chart">
      <row key="kinds" gap={2} y="center" role="radiogroup" label="Chart kind">
        {KINDS.map(option => (
          <button
            key={option.id}
            focusable={false}
            paddingLeft={7}
            paddingRight={7}
            paddingTop={3}
            paddingBottom={3}
            borderRadius={5}
            cursor="pointer"
            backgroundColor={kind.pipe(map(is => (is === option.id ? 'controlBackgroundHovered' : 'transparent')))}
            borderColor="controlBorder"
            borderWidth={1}
            role="radio"
            label={option.label}
            states={kind.pipe(map((is): readonly UiSemanticState[] => (is === option.id ? ['checked'] : [])))}
            onClick={() => pickKind(option.id)}>
            <text text={option.label} fontSize={11} textWrap="none" color="controlForeground" selectable={false} />
          </button>
        ))}
      </row>

      <editabletext
        key="title"
        ref={inputs.ref.value ?? undefined}
        value={title as never}
        width={180}
        fontSize={12}
        color="text"
        textWrap="none"
        verticalAlign="middle"
        backgroundColor="background"
        borderColor="border"
        borderWidth={1}
        padding={4}
        role="textbox"
        label="Chart title"
        onInput={(event: UiTextChangeEvent) => (title.value = event.value)}
        onKeyDown={onFieldKey(() => (editingId.value === 0 ? insert() : applyTitle()))}
      />

      {button('Insert', insert)}
      {button('Rename', applyTitle, hasChart)}
      {button('Delete', remove, hasChart)}
      <box flex={1} minWidth={0} />
      {button('Close', close)}
    </row>
  );
}
