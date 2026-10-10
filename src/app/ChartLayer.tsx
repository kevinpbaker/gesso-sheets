import { Box, Paint, type PaintBox, type PaintSurface, type UiElement, type UiPointerEvent } from 'gesso-core';

import { MIN_CHART_HEIGHT, MIN_CHART_WIDTH } from '../sheet/Chart';
import { drawChart } from './ChartPaint';
import { GUTTER_WIDTH } from './dimensions';
import type { SheetChart, SheetChartSeries } from './SheetContract';

/**
 * The floating object layer, which is one element and a handful of
 * handles.
 *
 * **A chart is a child of a row.** That is the whole trick, and it is
 * the one this file already uses for the fill handle and the formula
 * hint: a child of a row travels with it through a scroll and needs
 * nothing kept in step. The alternative — an overlay positioned from
 * the scroll offset — is a subscription to the scroll and a node
 * repositioned on every frame of one, to arrive at the same pixel.
 *
 * The row it hangs off is the row the chart *starts* in, not the
 * first row in view. That matters: hanging it off the window's first
 * row would unmount and rebuild the chart on every row scrolled past,
 * and a rebuilt `Paint` is a repainted picture. Anchored, the node
 * stays and only its row moves, which is layout and costs nothing.
 * Keeping the anchor row mounted while the chart is on screen is
 * `extendRange`'s job — the same hook a merge reaching up out of the
 * window uses, for the same reason.
 *
 * Nothing here decides what a chart looks like; `ChartPaint.ts` draws
 * and `Chart.ts` does the arithmetic. This is placement, selection
 * and the two drags.
 */

/** The size of a corner handle, and how far it hangs outside. */
const HANDLE = 9;

export interface ChartDrag {
  readonly id: number;
  /** What is being dragged: the whole chart, or one of its corners. */
  readonly corner: Corner | null;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export type Corner = 'topLeft' | 'topRight' | 'bottomLeft' | 'bottomRight';

export interface ChartCallbacks {
  readonly select: (id: number) => void;
  /**
   * A drag started, at a pointer position.
   *
   * The position rather than a delta, because a `UiPointerEvent`
   * carries where the pointer *is* and nothing about where it was —
   * which is the same reason the column resize remembers its own
   * starting x.
   */
  readonly begin: (id: number, corner: Corner | null, x: number, y: number) => void;
  readonly move: (x: number, y: number) => void;
  readonly end: () => void;
}

export interface ChartElementProps {
  readonly chart: SheetChart;
  readonly data: SheetChartSeries | null;
  readonly selected: boolean;
  /** Where the row this hangs off begins, in sheet pixels. */
  readonly rowOffset: number;
  /** The live rectangle while this chart is being dragged, or null. */
  readonly drag: ChartDrag | null;
  readonly on: ChartCallbacks;
  /** Whether the theme is dark, which the picture is made for; see `ChartInk`. */
  readonly dark: boolean;
}

/**
 * One chart, placed in its row.
 *
 * The rectangle drawn is the drag's while one is happening and the
 * document's the rest of the time, so a move follows the pointer at
 * frame rate without a round trip — the same trade a column resize
 * makes, and for the same reason: the application worker learns where
 * it ended up, not where it passed through.
 */
export function chartElement(props: ChartElementProps): UiElement {
  const { chart, data, selected, rowOffset, drag, on, dark } = props;
  const live = drag !== null && drag.id === chart.id ? drag : null;
  const x = live?.x ?? chart.x;
  const y = live?.y ?? chart.y;
  const width = live?.width ?? chart.width;
  const height = live?.height ?? chart.height;

  const children: UiElement[] = [];
  if (selected) {
    for (const corner of ['topLeft', 'topRight', 'bottomLeft', 'bottomRight'] as Corner[]) {
      children.push(handle(corner, width, height, on, chart.id));
    }
  }

  // `Paint` rather than `Box`: it is a Box the application draws into,
  // and children stack over whatever it drew — which is what makes the
  // handles children of the chart rather than a layer beside it.
  return Paint(
    {
      key: `chart-${chart.id}`,
      position: 'absolute',
      // The row's own coordinates start at its left edge, which is the
      // gutter — the same arithmetic the fill handle does.
      left: GUTTER_WIDTH + x,
      top: y - rowOffset,
      width,
      height,
      // Over the cells and under the chrome. A menu opened across a
      // chart has to cover it.
      zIndex: 5,
      cursor: 'grab',
      // `image` with a label, which is what a chart is to a screen
      // reader: a picture with a sentence describing it.
      role: 'image',
      label: chart.title === '' ? `Chart of ${chart.range}` : `${chart.title}, chart of ${chart.range}`,
      focusable: true,
      borderWidth: selected ? 1 : 0,
      borderColor: selected ? 'primary' : 'transparent',
      paint: {
        draw: (surface: PaintSurface, box: PaintBox) => drawChart(surface, box, chart, data, dark),
        /**
         * Everything the picture depends on, and nothing else.
         *
         * The engine repaints when one of these changes and at no
         * other time, so a chart standing still over a scrolling
         * sheet is one image draw a frame. `data` is compared by
         * identity, which is exactly right: the service rebuilds the
         * object when the numbers change and the differ hands the
         * same one back when they have not. And the theme: the
         * picture is made in it, and one made in the light theme was
         * shown on after View ▸ Theme went dark.
         */
        inputs: [chart.kind, chart.title, chart.legend, chart.range, width, height, data, dark]
      },
      onPointerDown: (event: UiPointerEvent) => {
        on.select(chart.id);
        // The grid sweeps a selection out of a pan, and a drag that
        // began on a chart is not sweeping cells.
        event.stopPropagation();
      },
      onPanStart: (event: UiPointerEvent) => {
        on.begin(chart.id, null, event.x, event.y);
        event.stopPropagation();
      },
      onPanMove: (event: UiPointerEvent) => {
        on.move(event.x, event.y);
        event.stopPropagation();
      },
      onPanEnd: (event: UiPointerEvent) => {
        on.end();
        event.stopPropagation();
      }
    },
    ...children
  );
}

function handle(corner: Corner, width: number, height: number, on: ChartCallbacks, id: number): UiElement {
  const left = corner === 'topLeft' || corner === 'bottomLeft';
  const top = corner === 'topLeft' || corner === 'topRight';
  return Box({
    key: corner,
    position: 'absolute',
    left: left ? -HANDLE / 2 : width - HANDLE / 2,
    top: top ? -HANDLE / 2 : height - HANDLE / 2,
    width: HANDLE,
    height: HANDLE,
    backgroundColor: 'primary',
    borderColor: 'background',
    borderWidth: 1,
    borderRadius: 2,
    zIndex: 6,
    cursor: left === top ? 'nwse-resize' : 'nesw-resize',
    role: 'button',
    label: `Resize ${id} from the ${spoken(corner)}`,
    onPointerDown: (event: UiPointerEvent) => event.stopPropagation(),
    onPanStart: (event: UiPointerEvent) => {
      on.begin(id, corner, event.x, event.y);
      event.stopPropagation();
    },
    onPanMove: (event: UiPointerEvent) => {
      on.move(event.x, event.y);
      event.stopPropagation();
    },
    onPanEnd: (event: UiPointerEvent) => {
      on.end();
      event.stopPropagation();
    }
  });
}

/**
 * A drag applied to a rectangle.
 *
 * Kept here rather than in the grid because it is arithmetic with two
 * awkward cases and no state: a corner drag moves the origin as well
 * as the size, and a chart squeezed past its minimum has to stop
 * growing on the *other* side rather than turning inside out.
 */
export function dragged(
  from: { x: number; y: number; width: number; height: number },
  corner: Corner | null,
  dx: number,
  dy: number
): { x: number; y: number; width: number; height: number } {
  if (corner === null) {
    return { ...from, x: Math.max(0, from.x + dx), y: Math.max(0, from.y + dy) };
  }
  const left = corner === 'topLeft' || corner === 'bottomLeft';
  const top = corner === 'topLeft' || corner === 'topRight';

  let { x, y, width, height } = from;
  if (left) {
    // Clamped before the origin moves, so a corner dragged past the
    // minimum pins the left edge instead of pushing the right one.
    const wanted = Math.max(MIN_CHART_WIDTH, from.width - dx);
    x = from.x + (from.width - wanted);
    width = wanted;
  } else {
    width = Math.max(MIN_CHART_WIDTH, from.width + dx);
  }
  if (top) {
    const wanted = Math.max(MIN_CHART_HEIGHT, from.height - dy);
    y = from.y + (from.height - wanted);
    height = wanted;
  } else {
    height = Math.max(MIN_CHART_HEIGHT, from.height + dy);
  }
  return { x: Math.max(0, x), y: Math.max(0, y), width, height };
}

function spoken(corner: Corner): string {
  return corner === 'topLeft'
    ? 'top left'
    : corner === 'topRight'
      ? 'top right'
      : corner === 'bottomLeft'
        ? 'bottom left'
        : 'bottom right';
}
