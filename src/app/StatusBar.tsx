import { map, type Observable } from 'rxjs';

import { percent } from 'gesso-core';
import { type ComponentContext, type Inputs } from 'gesso-framework';

import { Sheet, type SheetStatus } from './SheetContract';
import { describeStats } from './Statistics';

/**
 * The line along the bottom: what the selection adds up to, and what
 * the application thread still owes.
 *
 * Two readouts and not one, because they answer two different
 * questions and only one of them is about the document. The totals
 * are what a person came for; the pending count is the thing this
 * application exists to make visible, and moving it down here from
 * the formula bar puts it where every other spreadsheet keeps its
 * "Calculating…".
 */
export function StatusBar(_inputs: Inputs<{}>, ctx: ComponentContext) {
  const sheet = ctx.channel(Sheet);

  const totals: Observable<string> = sheet.view.stats.pipe(map(describeStats));

  /**
   * What the application thread is doing.
   *
   * The second half is what makes the first half worth reading.
   * "Ready" on its own is also what a sheet that did nothing would
   * say.
   */
  const work: Observable<string> = sheet.view.status.pipe(
    map((status: SheetStatus) => {
      const done = status.evaluated === 0 ? '' : ` · ${status.evaluated.toLocaleString('en-US')} evaluated`;
      return `${status.pending === 0 ? 'Ready' : `${status.pending.toLocaleString('en-US')} to do`}${done}`;
    })
  );

  /** The search, when there is one. Empty is the usual answer. */
  const search: Observable<string> = sheet.view.find.pipe(
    map(find => {
      if (find.query === '') {
        return '';
      }
      if (find.matches === 0) {
        return 'No matches';
      }
      return find.active === 0
        ? `${find.matches.toLocaleString('en-US')} matches`
        : `${find.active} of ${find.matches.toLocaleString('en-US')}`;
    })
  );

  return (
    <column width={percent(100)} flexShrink={0}>
      <box width={percent(100)} height={1} backgroundColor="border" />
      <row
        width={percent(100)}
        y="center"
        gap={16}
        paddingLeft={10}
        paddingRight={10}
        paddingTop={4}
        paddingBottom={4}
        backgroundColor="surface"
        role="status"
        label="Sheet status">
        <text text={totals} fontSize={11} color="text" verticalAlign="middle" selectable={false} />
        <box flex={1} minWidth={0} />
        <text text={search} fontSize={11} color="textMuted" verticalAlign="middle" selectable={false} />
        <text
          text={work}
          fontSize={11}
          color="textMuted"
          verticalAlign="middle"
          textAlign="end"
          selectable={false}
          live="polite"
        />
      </row>
    </column>
  );
}
