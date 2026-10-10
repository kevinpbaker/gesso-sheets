import { combineLatest, map } from 'rxjs';

import type { ComponentContext, Inputs } from 'gesso-framework';

import { Sheet, type SheetEditor, type SheetSimulation } from './SheetContract';
import type { SheetEditing } from './SheetEditing';

/**
 * What the last simulation said, beside the tabs.
 *
 * While one runs: how far it has got, and Stop. Once it has run: for
 * the cell you are on, the range its value fell in — one trial in ten
 * below P10, half below P50, nine in ten below P90 — and Clear. The
 * histograms in the cells say the shape; this says the numbers.
 *
 * Nothing at all while there is no simulation, so a workbook without
 * guesses in it shows the strip it always did.
 */
export interface SimulationReadoutProps {
  readonly editing: SheetEditing;
}

export function SimulationReadout(inputs: Inputs<SimulationReadoutProps>, ctx: ComponentContext) {
  const sheet = ctx.channel(Sheet);
  const edit = inputs.editing.value;

  const button = (key: string, text: string, label: string, onClick: () => void) => (
    <button
      key={key}
      onClick={() => {
        onClick();
        edit.focusSheet();
      }}
      label={label}
      paddingLeft={7}
      paddingRight={7}
      paddingTop={2}
      paddingBottom={2}
      borderRadius={4}
      backgroundColor="surface"
      borderColor="border"
      borderWidth={1}
      cursor="pointer">
      <text text={text} fontSize={11} color="text" selectable={false} />
    </button>
  );

  const line = (key: string, text: string) => (
    <text key={key} text={text} fontSize={11} color="textMuted" verticalAlign="middle" selectable={false} textWrap="none" live="polite" />
  );

  const parts = combineLatest([sheet.view.simulation, sheet.view.editor]).pipe(
    map(([run, editor]: [SheetSimulation, SheetEditor]) => {
      if (run.state === 'idle') {
        return [];
      }
      if (run.state === 'running') {
        return [
          line('progress', `Simulating ${run.on}: ${run.done.toLocaleString('en-US')} of ${run.trials.toLocaleString('en-US')} trials`),
          button('stop', 'Stop', 'Stop the simulation', () => sheet.send.stopSimulation())
        ];
      }
      const spread = editor.spread;
      const said =
        spread === null
          ? `${run.done.toLocaleString('en-US')} trials of ${run.on} · pick a cell with bars for its range`
          : `P10 ${spread.p10} · P50 ${spread.p50} · P90 ${spread.p90} · mean ${spread.mean}`;
      return [line('said', said), button('clear', 'Clear', 'Clear the simulation', () => sheet.send.stopSimulation())];
    })
  );

  return (
    <row gap={6} y="center" minWidth={0}>
      {parts}
    </row>
  );
}
