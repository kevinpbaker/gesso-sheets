import type { Observable } from 'rxjs';
import { distinctUntilChanged, map } from 'rxjs/operators';

import { LazySheet, percent, Row, Text, type UiElement, type UiVirtualSheet } from 'gesso-core';
import { FrameService, internalState, type ComponentContext, type Inputs } from 'gesso-framework';

import { BENCH_DONE, BENCH_PREFIX, BenchDriver, benchMatrix, isBench } from './spike/Bench';
import {
  COLUMN_COUNT,
  ROW_COUNT,
  Sheet,
  valueAt,
  type SheetBlock,
  type SheetStats,
  type WireShape
} from './spike/SheetContract';

/**
 * Phase 0, on the screen.
 *
 * A million cells, no formulas, and every number this phase has to
 * produce shown while it scrolls. Nothing here is meant to survive:
 * the point is to find out whether the window can be windowed on both
 * axes and whether the round trip to the application worker keeps up,
 * and then to throw it away.
 *
 * What the screen has to make checkable:
 *
 *   - **60fps on both axes.** `Sweep ↓` and `Sweep →` scroll at a
 *     fixed velocity for a fixed time, so a run is repeatable and the
 *     frame gap is not a measure of how fast someone moved a finger.
 *   - **The band.** The two bands are separate controls because they
 *     are not the same purchase: the mount band costs nodes on every
 *     frame, the fetch band costs only cells on the wire. Which one
 *     buys coverage more cheaply is the thing to find out.
 *   - **The round trip.** `Missed` is the share of visible cells the
 *     application worker had not sent yet, counted every frame. That
 *     number going to zero is what "the band is wide enough" means.
 */

const ROW_HEIGHT = 24;
const COLUMN_WIDTH = 104;

/** How fast a sweep scrolls, in logical pixels per second. */
const SPEEDS = [600, 3000, 9000];
const SWEEP_MS = 4000;

interface Readout {
  fps: number;
  frameMs: number;
  worstFrameMs: number;
  measured: number;
  nodes: number;
  frames: number;
  missFrames: number;
  worstMiss: number;
  /** The mounted range, so a screenshot says where the sheet is. */
  window: string;
}

const IDLE: Readout = {
  fps: 0,
  frameMs: 0,
  worstFrameMs: 0,
  measured: 0,
  nodes: 0,
  frames: 0,
  missFrames: 0,
  worstMiss: 0,
  window: '—'
};

export function App(_inputs: Inputs<{}>, ctx: ComponentContext) {
  const sheet = ctx.channel(Sheet);
  const block = sheet.view.block;

  const mountBandRows = internalState(2);
  const mountBandColumns = internalState(1);
  const fetchBandRows = internalState(0);
  const fetchBandColumns = internalState(0);
  const speed = internalState(SPEEDS[1]);
  const shape = internalState<WireShape>('keyed');
  const busy = internalState(0);
  const readout = internalState<Readout>(IDLE);
  const scrollX = internalState(0);
  const scrollY = internalState(0);

  /**
   * The cell elements now in the window, by the cell they hold.
   *
   * The first run of the bench said the horizontal axis cost 10 ms a
   * frame in the windowing phase against 2 ms for the vertical one,
   * with a lower frame cost and a smaller wire than the vertical runs
   * — so it was neither layout nor the barrier. It was this: a column
   * window that moves invalidates every mounted row, and rebuilding a
   * row rebuilt its cells, and building a cell allocates two RxJS
   * pipelines. Around eleven hundred of them, one or two times a
   * frame.
   *
   * An element is immutable, so the cells that are still in the window
   * can simply be handed back. What is left to build on a column step
   * is the one column that entered.
   */
  const cells = new Map<string, UiElement>();

  /**
   * One cell.
   *
   * Two bindings and no state: the node is keyed by the cell it holds,
   * so it lives exactly as long as that cell is in the window and its
   * text is written rather than rebuilt when a value arrives. `null`
   * is the cell the application worker has not sent yet — drawn as a
   * dot in the placeholder colour rather than left blank, so a gap
   * reads as "not here yet" instead of as an empty sheet.
   */
  const buildCell = (row: number, column: number): UiElement => {
    const value = block.pipe(
      map(current => valueAt(current, row, column)),
      distinctUntilChanged()
    );
    return Text({
      key: column,
      text: value.pipe(map(text => text ?? '·')),
      color: value.pipe(map(text => (text === null ? 'placeholder' : column < 2 ? 'textMuted' : 'text'))),
      width: COLUMN_WIDTH,
      height: ROW_HEIGHT,
      flexShrink: 0,
      paddingLeft: 8,
      paddingRight: 8,
      fontSize: 12,
      textWrap: 'none',
      textOverflow: 'clip',
      verticalAlign: 'middle',
      textAlign: column < 2 ? 'start' : 'right',
      backgroundColor: row % 2 === 0 ? 'background' : 'surface'
    });
  };

  const cell = (row: number, column: number): UiElement => {
    const key = `${row}:${column}`;
    let built = cells.get(key);
    if (built === undefined) {
      built = buildCell(row, column);
      cells.set(key, built);
    }
    return built;
  };

  const renderRow = (row: number, firstColumn: number, lastColumn: number): UiElement => {
    const line: UiElement[] = [];
    for (let column = firstColumn; column <= lastColumn; column++) {
      line.push(cell(row, column));
    }
    return Row({}, ...line);
  };

  let window: UiVirtualSheet | undefined;
  const grid = LazySheet(
    {
      flex: 1,
      minHeight: 0,
      width: percent(100),
      backgroundColor: 'background',
      rowCount: ROW_COUNT,
      columnCount: COLUMN_COUNT,
      rowHeight: ROW_HEIGHT,
      columnWidth: COLUMN_WIDTH,
      rowOverscan: mountBandRows,
      columnOverscan: mountBandColumns,
      scrollX,
      scrollY,
      sheetRef: found => (window = found)
    },
    renderRow
  );
  if (window === undefined) {
    throw new Error('LazySheet did not hand back its window.');
  }
  const sheetWindow = window;

  // The cache is bounded by the window: anything outside it is a cell
  // that scrolled away, and holding it would hold its two
  // subscriptions with it.
  ctx.effect(sheetWindow.range$, range => {
    for (const key of cells.keys()) {
      const colon = key.indexOf(':');
      const row = Number(key.slice(0, colon));
      const column = Number(key.slice(colon + 1));
      if (row < range.firstRow || row > range.lastRow || column < range.firstColumn || column > range.lastColumn) {
        cells.delete(key);
      }
    }
  });

  // The round trip itself: the range the window settled on this frame
  // is what the application worker is asked for. `range$` emits only
  // when the range changes, so this is not a command per frame.
  ctx.effect(sheetWindow.range$, range =>
    sheet.send.setViewport(range.firstRow, range.lastRow, range.firstColumn, range.lastColumn)
  );
  ctx.effect(fetchBandRows, rows => sheet.send.setFetchBand(rows, fetchBandColumns.value));
  ctx.effect(fetchBandColumns, columns => sheet.send.setFetchBand(fetchBandRows.value, columns));
  ctx.effect(shape, next => sheet.send.setShape(next));
  ctx.effect(busy, ms => sheet.send.setBusy(ms));

  // ---------------------------------------------------------------------
  // Measurement
  // ---------------------------------------------------------------------

  let sweep: { axis: 'x' | 'y'; until: number; last: number } | null = null;
  let stats = { ...IDLE };
  let lastAt = 0;

  const reset = (): void => {
    stats = { ...IDLE };
    lastAt = 0;
    readout.value = IDLE;
  };

  /**
   * The unattended run, when the page url carries `?bench`.
   *
   * It drives the same knobs the buttons do, from the same frame
   * stream, so what it measures is what the screen does.
   */
  const bench = isBench()
    ? new BenchDriver(
        benchMatrix(),
        {
          setMountBand: (rows, columns) => {
            mountBandRows.value = rows;
            mountBandColumns.value = columns;
          },
          setFetchBand: (rows, columns) => {
            fetchBandRows.value = rows;
            fetchBandColumns.value = columns;
          },
          setShape: next => (shape.value = next),
          setBusy: ms => (busy.value = ms),
          scrollTo: (x, y) => {
            scrollX.value = x;
            scrollY.value = y;
          },
          report: line => console.log(BENCH_PREFIX + line),
          done: () => console.log(BENCH_DONE)
        },
        { rowHeight: ROW_HEIGHT, columnWidth: COLUMN_WIDTH, rowCount: ROW_COUNT, columnCount: COLUMN_COUNT }
      )
    : null;

  ctx.effect(ctx.inject(FrameService).frames, frame => {
    // `at` is the render worker's own clock. Gaps between consecutive
    // values are the honest frame time; arrival times are not, because
    // a blocked main thread delivers a burst of them at once.
    const lastGap = lastAt > 0 ? frame.at - lastAt : 0;
    if (lastGap > 0) {
      stats.frames++;
      stats.frameMs += (lastGap - stats.frameMs) / Math.min(stats.frames, 30);
      stats.worstFrameMs = Math.max(stats.worstFrameMs, lastGap);
    }
    lastAt = frame.at;
    stats.measured = frame.measured;
    stats.nodes = frame.nodes;

    const missed = missing(sheetWindow.range$.value, block.value, mountBandRows.value, mountBandColumns.value);
    if (missed > 0) {
      stats.missFrames++;
      stats.worstMiss = Math.max(stats.worstMiss, missed);
    }

    if (sweep !== null) {
      const step = (speed.value * (frame.at - sweep.last)) / 1000;
      sweep.last = frame.at;
      if (sweep.axis === 'y') {
        scrollY.value = Math.min(ROW_COUNT * ROW_HEIGHT, scrollY.value + step);
      } else {
        scrollX.value = Math.min(COLUMN_COUNT * COLUMN_WIDTH, scrollX.value + step);
      }
      if (frame.at >= sweep.until) {
        sweep = null;
      }
    }

    const range = sheetWindow.range$.value;
    readout.value = {
      ...stats,
      fps: stats.frameMs > 0 ? 1000 / stats.frameMs : 0,
      window: `r${range.firstRow}–${range.lastRow} × c${range.firstColumn}–${range.lastColumn}`
    };

    const wire = sheet.view.stats.value;
    bench?.frame(frame.at, {
      gapMs: lastGap,
      durationMs: frame.durationMs,
      phases: frame.phases,
      missed,
      nodes: frame.nodes,
      measured: frame.measured,
      publishes: wire.publishes,
      patches: wire.patches,
      bytes: wire.bytes,
      cells: wire.lastCells
    });
  });

  const startSweep = (axis: 'x' | 'y'): void => {
    reset();
    scrollX.value = axis === 'x' ? 0 : scrollX.value;
    scrollY.value = axis === 'y' ? 0 : scrollY.value;
    const now = performance.now();
    sweep = { axis, until: now + SWEEP_MS, last: now };
  };

  return (
    <column width={percent(100)} height={percent(100)} backgroundColor="background">
      {Hud({
        readout,
        stats: sheet.view.stats,
        mountBandRows,
        mountBandColumns,
        fetchBandRows,
        fetchBandColumns,
        speed,
        shape,
        busy,
        onSweepDown: () => startSweep('y'),
        onSweepRight: () => startSweep('x'),
        onReset: reset
      })}
      {grid}
    </column>
  );
}

/**
 * The share of *visible* cells the application worker has not sent.
 *
 * The mounted range is inset by the mount band first, because a band
 * cell holding nothing is the band doing its job — it is off screen.
 * What matters is a cell the eye can reach that has no value in it.
 */
function missing(
  range: { firstRow: number; lastRow: number; firstColumn: number; lastColumn: number },
  block: SheetBlock,
  bandRows: number,
  bandColumns: number
): number {
  const firstRow = range.firstRow + (range.firstRow > 0 ? bandRows : 0);
  const lastRow = range.lastRow - bandRows;
  const firstColumn = range.firstColumn + (range.firstColumn > 0 ? bandColumns : 0);
  const lastColumn = range.lastColumn - bandColumns;
  const rows = lastRow - firstRow + 1;
  const columns = lastColumn - firstColumn + 1;
  if (rows <= 0 || columns <= 0) {
    return 0;
  }
  const covered =
    Math.max(0, Math.min(lastRow, block.lastRow) - Math.max(firstRow, block.firstRow) + 1) *
    Math.max(0, Math.min(lastColumn, block.lastColumn) - Math.max(firstColumn, block.firstColumn) + 1);
  return (rows * columns - covered) / (rows * columns);
}

// ---------------------------------------------------------------------------
// The readout
// ---------------------------------------------------------------------------

interface HudInputs {
  readout: ReturnType<typeof internalState<Readout>>;
  stats: Observable<SheetStats>;
  mountBandRows: ReturnType<typeof internalState<number>>;
  mountBandColumns: ReturnType<typeof internalState<number>>;
  fetchBandRows: ReturnType<typeof internalState<number>>;
  fetchBandColumns: ReturnType<typeof internalState<number>>;
  speed: ReturnType<typeof internalState<number>>;
  shape: ReturnType<typeof internalState<WireShape>>;
  busy: ReturnType<typeof internalState<number>>;
  onSweepDown: () => void;
  onSweepRight: () => void;
  onReset: () => void;
}

function Hud(inputs: HudInputs) {
  const { readout, stats } = inputs;
  return (
    <column
      width={percent(100)}
      flexShrink={0}
      gap={6}
      padding={10}
      backgroundColor="surface"
      borderColor="border"
      borderWidth={1}>
      <row gap={16} y="center" flexWrap="wrap">
        {stat('FPS', readout.pipe(map(r => (r.fps === 0 ? '—' : r.fps.toFixed(0)))))}
        {stat('Frame', readout.pipe(map(r => `${r.frameMs.toFixed(1)} ms`)))}
        {stat('Worst', readout.pipe(map(r => `${r.worstFrameMs.toFixed(0)} ms`)))}
        {stat('Window', readout.pipe(map(r => r.window)))}
        {stat('Nodes', readout.pipe(map(r => String(r.nodes))))}
        {stat('Measured', readout.pipe(map(r => String(r.measured))))}
        {stat(
          'Missed',
          readout.pipe(
            map(r =>
              r.frames === 0
                ? '—'
                : `${((100 * r.missFrames) / r.frames).toFixed(0)}% of frames, worst ${(100 * r.worstMiss).toFixed(0)}%`
            )
          )
        )}
      </row>
      <row gap={16} y="center" flexWrap="wrap">
        {stat('Sent', stats.pipe(map(s => `${s.lastCells} cells`)))}
        {stat('Patches', stats.pipe(map(s => String(s.lastPatches))))}
        {stat('Bytes', stats.pipe(map(s => `${(s.lastBytes / 1024).toFixed(1)} KiB`)))}
        {stat('Publishes', stats.pipe(map(s => String(s.publishes))))}
      </row>
      <row gap={8} y="center" flexWrap="wrap">
        {counter('Mount band', inputs.mountBandRows, inputs.mountBandColumns)}
        {counter('Fetch band', inputs.fetchBandRows, inputs.fetchBandColumns)}
        {button('Sweep ↓', inputs.onSweepDown)}
        {button('Sweep →', inputs.onSweepRight)}
        {button('Reset', inputs.onReset)}
        {cycle('Speed', inputs.speed, SPEEDS, value => `${value} px/s`)}
        {cycle('Wire', inputs.shape, ['keyed', 'columnKeyed', 'rows'] as WireShape[], value => value)}
        {cycle('Recalc', inputs.busy, [0, 8, 30], value => `${value} ms`)}
      </row>
    </column>
  );
}

function stat(label: string, value: Observable<string>) {
  return (
    <row gap={6} y="center">
      <text text={label} fontSize={11} color="textMuted" />
      <text text={value} fontSize={13} fontWeight={600} color="text" />
    </row>
  );
}

function button(label: string, onClick: () => void) {
  return (
    <button
      onClick={onClick}
      paddingLeft={10}
      paddingRight={10}
      paddingTop={5}
      paddingBottom={5}
      borderRadius={6}
      backgroundColor="controlBackground"
      borderColor="controlBorder"
      borderWidth={1}
      cursor="pointer">
      <text text={label} fontSize={12} color="controlForeground" />
    </button>
  );
}

/** A pair of counters, because a band has a row half and a column half. */
function counter(
  label: string,
  rows: ReturnType<typeof internalState<number>>,
  columns: ReturnType<typeof internalState<number>>
) {
  const step = (cell: ReturnType<typeof internalState<number>>, by: number) => () => {
    cell.value = Math.max(0, cell.value + by);
  };
  return (
    <row gap={4} y="center">
      <text text={label} fontSize={11} color="textMuted" />
      {button('−', step(rows, -1))}
      <text text={rows.pipe(map(value => `${value} r`))} fontSize={12} color="text" />
      {button('+', step(rows, 1))}
      {button('−', step(columns, -1))}
      <text text={columns.pipe(map(value => `${value} c`))} fontSize={12} color="text" />
      {button('+', step(columns, 1))}
    </row>
  );
}

/** A control that walks a short list, which is every remaining knob. */
function cycle<T>(
  label: string,
  cell: ReturnType<typeof internalState<T>>,
  values: readonly T[],
  format: (value: T) => string
) {
  const next = () => {
    const at = values.indexOf(cell.value);
    cell.value = values[(at + 1) % values.length];
  };
  return (
    <row gap={4} y="center">
      <text text={label} fontSize={11} color="textMuted" />
      <button
        onClick={next}
        paddingLeft={10}
        paddingRight={10}
        paddingTop={5}
        paddingBottom={5}
        borderRadius={6}
        backgroundColor="controlBackground"
        borderColor="controlBorder"
        borderWidth={1}
        cursor="pointer">
        <text text={cell.pipe(map(format))} fontSize={12} color="controlForeground" />
      </button>
    </row>
  );
}
