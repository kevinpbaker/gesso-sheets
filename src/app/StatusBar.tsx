import { BehaviorSubject, combineLatest, map, type Observable } from 'rxjs';

import { percent, type UiPointerEvent } from 'gesso-core';
import { Menu, type MenuItem } from 'gesso-components';
import { createComponent, internalState, persisted, ShellService, ShellStorage, type ComponentContext, type Inputs } from 'gesso-framework';

import { Sheet, zoomStep, type SheetStatus } from './SheetContract';
import { DEFAULT_FIGURES, FIGURES, figuresOf, isFigure, type StatFigure } from './Statistics';

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
  const shell = ctx.inject(ShellService);

  /**
   * Which figures the readout shows: the person's, and kept for them.
   *
   * In the browser's `localStorage`, through the shell — the render
   * worker cannot reach it, and `ShellStorage` is the engine's way
   * round that — because a preference is the one thing that belongs
   * there: small, the person's rather than the document's, and wanted
   * before the first frame of the next visit. Excel keeps this choice
   * for the whole application, and so does this. Where the browser
   * will not store anything, the choice lasts for the session and the
   * default comes back, which is all the storage could promise.
   */
  const figures = persisted(new ShellStorage(shell, { prefix: 'gessosheet:' }), 'status-figures', {
    initial: DEFAULT_FIGURES,
    revive: raw => (Array.isArray(raw) && raw.every(isFigure) ? (raw as StatFigure[]) : null),
    label: 'status bar figures'
  });
  ctx.onUnmount(() => figures.dispose());
  const chosen = figures.value;
  const copiedNote = internalState('');
  let noteTimer: ReturnType<typeof setTimeout> | null = null;
  const copy = (label: string, text: string): void => {
    shell.copyText(text);
    copiedNote.value = `${label} copied`;
    if (noteTimer !== null) {
      clearTimeout(noteTimer);
    }
    noteTimer = setTimeout(() => (copiedNote.value = ''), 2000);
  };

  /**
   * The figures, one button each: a click copies the figure, whole,
   * because the number somebody checked is usually the number they
   * wanted to put somewhere.
   */
  const totals = combineLatest([sheet.view.stats, chosen]).pipe(
    map(([stats, which]) =>
      figuresOf(stats, which).map(figure => (
        <text
          key={figure.id}
          text={`${figure.label} ${figure.value}`}
          fontSize={11}
          color="text"
          verticalAlign="middle"
          selectable={false}
          role="button"
          label={`${figure.label} ${figure.value}`}
          cursor="pointer"
          onClick={() => copy(figure.label, figure.copy)}
        />
      ))
    )
  );

  /**
   * The menu of figures, from a right-click on the readout or the
   * button beside it — the button is there so a keyboard can reach
   * the choice, and a pointer that does not think to right-click can.
   */
  const menuOpen = new BehaviorSubject(false);
  const menuAt = new BehaviorSubject({ x: 0, y: 0 });
  const menuItems = chosen.pipe(
    map((which): readonly MenuItem[] =>
      FIGURES.map(figure => ({ value: figure.id, label: `${which.includes(figure.id) ? '✓' : '  '} ${figure.label}` }))
    )
  );
  const openMenu = (at: { x: number; y: number }): void => {
    menuAt.next(at);
    menuOpen.next(true);
  };
  /** Where the chooser is, so the menu opens beside it and not wherever a key was pressed. */
  const chooser = ctx.bounds('chooser');
  const openAtChooser = (): void => {
    const box = chooser.value;
    openMenu({ x: box.x, y: box.y });
  };
  const menu = createComponent(Menu, {
    open: menuOpen,
    onOpenChange: (open: boolean) => menuOpen.next(open),
    items: menuItems,
    at: menuAt,
    label: 'Status bar figures',
    onSelect: (value: string) => {
      menuOpen.next(false);
      const id = value as StatFigure;
      const now = chosen.value;
      figures.set(now.includes(id) ? now.filter(figure => figure !== id) : [...now, id]);
    }
  });

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

  /**
   * Which document this is, and whether it has changes its file does
   * not. A browser tab's title is the shell's and not this thread's to
   * write, so the name goes where the rest of the document's state
   * already is.
   */
  const document: Observable<string> = sheet.view.document.pipe(
    map(current => {
      if (current.id === '') {
        return '';
      }
      const name = current.file?.name ?? current.name;
      if (current.elsewhere) {
        return `${name} · open in another tab, not saved here`;
      }
      return current.edited && current.file !== null ? `${name} · edited` : name;
    })
  );

  /** What the last file dropped on the window became, until the next edit. */
  const report: Observable<string> = sheet.view.transfer.pipe(map(transfer => transfer.report));

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
        <text text={document} fontSize={11} color="textMuted" verticalAlign="middle" selectable={false} />
        <row
          gap={12}
          y="center"
          role="group"
          label="Selection figures"
          onContextMenu={(event: UiPointerEvent) => openMenu({ x: event.x, y: event.y })}>
          {totals}
        </row>
        <text
          text="▾"
          fontSize={11}
          color="textMuted"
          verticalAlign="middle"
          selectable={false}
          role="button"
          label="Choose what the status bar shows"
          cursor="pointer"
          focusable={true}
          modifiers={[chooser.modifier]}
          onClick={openAtChooser}
          onKeyDown={event => {
            if (event.key === 'Enter' || event.key === ' ') {
              event.preventDefault();
              openAtChooser();
            }
          }}
        />
        <text text={copiedNote} fontSize={11} color="textMuted" verticalAlign="middle" selectable={false} live="polite" />
        {menu}
        <box flex={1} minWidth={0} />
        <text
          text={report}
          fontSize={11}
          color="textMuted"
          verticalAlign="middle"
          selectable={false}
          live="polite"
        />
        <text text={search} fontSize={11} color="textMuted" verticalAlign="middle" selectable={false} />
        {zoomControl('Zoom out', '−', () => sheet.send.setZoom(zoomStep(sheet.view.geometry.value.zoom, 'zoomOut')))}
        {zoomControl(
          'Zoom to 100%',
          sheet.view.geometry.pipe(map(geometry => `${Math.round(geometry.zoom * 100)}%`)),
          () => sheet.send.setZoom(1)
        )}
        {zoomControl('Zoom in', '+', () => sheet.send.setZoom(zoomStep(sheet.view.geometry.value.zoom, 'zoomIn')))}
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

/**
 * One of the zoom controls at the end of the bar: −, the zoom itself,
 * and +. The zoom's own figure goes back to 100% on a click, which is
 * the question somebody clicking a percentage is usually asking.
 */
function zoomControl(label: string, text: string | Observable<string>, onClick: () => void) {
  return (
    <text
      text={text}
      fontSize={11}
      color="textMuted"
      verticalAlign="middle"
      selectable={false}
      role="button"
      label={label}
      cursor="pointer"
      onClick={onClick}
    />
  );
}
