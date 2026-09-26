import { map } from 'rxjs';

import { percent, type UiKeyboardEvent, type UiNode } from 'gesso-core';
import { type ComponentContext, type Inputs } from 'gesso-framework';

import type { FileActions } from './Files';

/**
 * The files the shell remembers, as a row of buttons.
 *
 * A bar rather than a submenu, because the menus are a table fixed at
 * build time and this list is the shell's, read when the bar opens. A
 * row in the flow for the reason the find bar is one: it covers
 * nothing, and a person can see the sheet they are about to leave.
 *
 * Choosing a file is a click, which is also what the browser needs to
 * ask for permission to read it again after a reload — so the request
 * goes from this click straight to the shell, before anything else.
 */

export interface RecentBarProps {
  readonly files: FileActions;
  readonly onClose: () => void;
  readonly ref?: (node: UiNode | null) => void;
}

export function RecentBar(inputs: Inputs<RecentBarProps>, _ctx: ComponentContext) {
  const files = inputs.files.value;
  files.refreshRecent();

  const close = (): void => inputs.onClose.value();

  const onKey = (event: UiKeyboardEvent): void => {
    if (event.key === 'Escape') {
      event.preventDefault();
      event.stopPropagation();
      close();
    }
  };

  const button = (key: string, label: string, onClick: () => void, ref?: (node: UiNode | null) => void) => (
    <button
      key={key}
      ref={ref}
      onClick={onClick}
      label={label}
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

  const list = files.recent.pipe(
    map(recent =>
      recent.length === 0
        ? [
            <text
              key="none"
              text="No files opened or saved yet."
              fontSize={11}
              color="textMuted"
              verticalAlign="middle"
              selectable={false}
            />
          ]
        : recent.slice(0, 8).map(file =>
            button(`file-${file.handle}`, file.name, () => {
              files.reopen(file.handle);
              close();
            })
          )
    )
  );

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
      label="Recent files"
      onKeyDown={onKey}>
      <text text="Recent" fontSize={11} color="textMuted" verticalAlign="middle" selectable={false} />
      {list}
      <box flex={1} minWidth={0} />
      {button('close', 'Close', close, inputs.ref.value ?? undefined)}
    </row>
  );
}
