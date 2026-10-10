import { combineLatest, map, type Observable } from 'rxjs';

import { Column, Text, type UiKeyboardEvent, type UiNode, type UiTextChangeEvent } from 'gesso-core';
import { Dialog } from 'gesso-components';
import { FocusService, internalState, type ComponentContext, type Inputs, type ShellRecentFile } from 'gesso-framework';

import type { SheetLibraryEntry, SheetLibraryView } from './SheetContract';
import { matching, whenSaid } from './whereItIs';

/**
 * Every workbook this browser keeps, to open, start, rename, copy and
 * delete — Phase 36's home screen.
 *
 * A dialog over the sheet rather than a page of its own: the sheet
 * behind it is the workbook somebody was in, and Escape puts them back
 * in it. One row per workbook, newest edit first, with the name as
 * the button that opens it and the three things that can be done to
 * it beside it. Deleting asks in the row itself, because a second
 * dialog over this one is a dialog people click through.
 *
 * Below the list, the files the shell remembers — *Open recent*'s
 * list, read from the same place — since a workbook saved to a file is
 * also a thing somebody comes back to, and the sentence at the bottom
 * says where each kind is kept.
 */
export interface HomeDialogProps {
  readonly open: boolean;
  readonly library: SheetLibraryView;
  readonly recent: readonly ShellRecentFile[];
  /** The clock "edited 5 minutes ago" is read from; the wall clock unless a spec says. */
  readonly now?: () => number;
  readonly onOpen: (id: string) => void;
  readonly onNew: () => void;
  readonly onOpenFile: () => void;
  readonly onReopen: (handle: number) => void;
  readonly onRename: (id: string, name: string) => void;
  readonly onDuplicate: (id: string) => void;
  readonly onDelete: (id: string) => void;
  readonly onClose: () => void;
}

/** What a row is doing: being read, renamed, or asked about deleting. */
type RowMode = { readonly id: string; readonly kind: 'rename' | 'delete' } | null;

export function HomeDialog(inputs: Inputs<HomeDialogProps>, ctx: ComponentContext) {
  const focus = ctx.inject(FocusService);
  const search = internalState('');
  const mode = internalState<RowMode>(null);
  let renamed = '';
  const now = (): number => (inputs.now.value ?? Date.now)();

  ctx.effect(inputs.open, open => {
    if (open) {
      search.value = '';
      mode.value = null;
    }
  });

  const button = (key: string, label: string, onClick: () => void, tone: 'plain' | 'primary' | 'danger' = 'plain', name = label) => (
    <button
      key={key}
      onClick={onClick}
      label={name}
      paddingLeft={10}
      paddingRight={10}
      paddingTop={4}
      paddingBottom={4}
      borderRadius={6}
      backgroundColor={tone === 'primary' ? 'primary' : tone === 'danger' ? 'danger' : 'controlBackground'}
      borderColor="controlBorder"
      borderWidth={1}
      cursor="pointer">
      <text
        text={label}
        fontSize={12}
        textWrap="none"
        color={tone === 'plain' ? 'controlForeground' : 'primaryForeground'}
        selectable={false}
      />
    </button>
  );

  const focusOnArrival = (node: UiNode | null): void => {
    if (node !== null) {
      queueMicrotask(() => focus.focus(node));
    }
  };

  const row = (entry: SheetLibraryEntry, current: RowMode) => {
    if (current?.id === entry.id && current.kind === 'delete') {
      return (
        <row key={`row-${entry.id}`} gap={8} y="center" minWidth={0} padding={4} backgroundColor="surface" borderRadius={6}>
          <text
            text={`Delete “${entry.name}”? Its versions go with it, and it cannot be undone.`}
            flex={1}
            minWidth={0}
            fontSize={12}
            color="text"
            textWrap="word"
          />
          {button(`confirm-${entry.id}`, 'Delete', () => {
            mode.value = null;
            inputs.onDelete.value(entry.id);
          }, 'danger', `Delete ${entry.name}`)}
          {button(`cancel-${entry.id}`, 'Keep it', () => (mode.value = null))}
        </row>
      );
    }
    if (current?.id === entry.id && current.kind === 'rename') {
      renamed = entry.name;
      const save = (): void => {
        mode.value = null;
        inputs.onRename.value(entry.id, renamed);
      };
      return (
        <row key={`row-${entry.id}`} gap={8} y="center" minWidth={0} padding={4}>
          <editabletext
            value={entry.name}
            ref={focusOnArrival}
            flex={1}
            minWidth={0}
            fontSize={13}
            color="text"
            textWrap="none"
            backgroundColor="background"
            borderColor="focusRing"
            borderWidth={1}
            padding={4}
            role="textbox"
            label={`New name for ${entry.name}`}
            onInput={(event: UiTextChangeEvent) => (renamed = event.value)}
            onKeyDown={(event: UiKeyboardEvent) => {
              if (event.key === 'Enter') {
                event.preventDefault();
                event.stopPropagation();
                save();
              } else if (event.key === 'Escape') {
                event.preventDefault();
                event.stopPropagation();
                mode.value = null;
              }
            }}
          />
          {button(`save-${entry.id}`, 'Save', save, 'primary', 'Save name')}
          {button(`cancel-${entry.id}`, 'Cancel', () => (mode.value = null))}
        </row>
      );
    }
    const detail = [whenSaid(entry.edited, now()), entry.file === '' ? 'in this browser' : `also ${entry.file}`]
      .concat(entry.open ? ['open now'] : [])
      .join(' · ');
    return (
      <row key={`row-${entry.id}`} gap={8} y="center" minWidth={0} padding={4}>
        <button
          onClick={() => inputs.onOpen.value(entry.id)}
          label={`Open ${entry.name}`}
          x="start"
          flex={1}
          minWidth={0}
          paddingLeft={8}
          paddingRight={8}
          paddingTop={5}
          paddingBottom={5}
          borderRadius={6}
          backgroundColor={entry.open ? 'controlBackgroundHovered' : 'controlBackground'}
          borderColor="controlBorder"
          borderWidth={1}
          cursor="pointer">
          <column gap={1} minWidth={0}>
            <text text={entry.name} fontSize={13} fontWeight="bold" color="controlForeground" textWrap="none" selectable={false} />
            <text text={detail} fontSize={11} color="textMuted" textWrap="none" selectable={false} />
          </column>
        </button>
        {button(`rename-${entry.id}`, 'Rename', () => (mode.value = { id: entry.id, kind: 'rename' }), 'plain', `Rename ${entry.name}`)}
        {button(`copy-${entry.id}`, 'Duplicate', () => inputs.onDuplicate.value(entry.id), 'plain', `Duplicate ${entry.name}`)}
        {button(`delete-${entry.id}`, 'Delete', () => (mode.value = { id: entry.id, kind: 'delete' }), 'plain', `Delete ${entry.name}`)}
      </row>
    );
  };

  const list = combineLatest([inputs.library, search, mode]).pipe(
    map(([library, words, current]) => {
      const shown = matching(library.entries, words);
      if (shown.length === 0) {
        return [
          <text
            key="none"
            text={library.entries.length === 0 ? 'No workbooks yet.' : `No workbook is called anything like “${words}”.`}
            fontSize={12}
            color="textMuted"
            textWrap="word"
          />
        ];
      }
      return shown.map(entry => row(entry, current));
    })
  );

  const recent: Observable<readonly ReturnType<typeof button>[]> = inputs.recent.pipe(
    map(files =>
      files.slice(0, 6).map(file => button(`file-${file.handle}`, file.name, () => inputs.onReopen.value(file.handle), 'plain', `Reopen ${file.name}`))
    )
  );

  return (
    <Dialog
      open={inputs.open}
      onClose={() => inputs.onClose.value()}
      title="Workbooks"
      width={640}
      content={Column(
        { gap: 10, minWidth: 0 },
        <row gap={8} y="center">
          <editabletext
            value={search}
            ref={(node: UiNode | null) => {
              if (inputs.open.value) {
                focusOnArrival(node);
              }
            }}
            flex={1}
            minWidth={0}
            fontSize={13}
            color="text"
            textWrap="none"
            backgroundColor="background"
            borderColor="border"
            borderWidth={1}
            padding={5}
            role="searchbox"
            label="Search workbooks"
            placeholder="Search workbooks"
            onInput={(event: UiTextChangeEvent) => (search.value = event.value)}
          />
          {button('new', 'New blank workbook', () => inputs.onNew.value(), 'primary')}
        </row>,
        <column gap={2} minWidth={0} maxHeight={320} overflow="auto" role="list" label="Workbooks in this browser">
          {list}
        </column>,
        <row gap={8} y="center" minWidth={0}>
          <text text="Files" fontSize={11} color="textMuted" selectable={false} />
          <row gap={6} y="center" flex={1} minWidth={0} overflow="hidden">
            {recent}
          </row>
          {button('openFile', 'Open a file…', () => inputs.onOpenFile.value())}
        </row>,
        Text({
          text:
            'Workbooks are kept in this browser’s private storage for this site, saved as you type. Clearing the site’s data deletes them. Save, or File ▸ Keep saving to this file, to keep a copy as a file on this computer.',
          fontSize: 11,
          color: 'textMuted',
          textWrap: 'word'
        }),
        <row gap={8} x="end">
          {button('close', 'Close', () => inputs.onClose.value())}
        </row>
      )}
    />
  );
}
