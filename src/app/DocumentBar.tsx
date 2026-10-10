import { distinctUntilChanged, map } from 'rxjs';

import { percent, type UiKeyboardEvent, type UiNode, type UiTextChangeEvent } from 'gesso-core';
import { Icon, tooltip } from 'gesso-components';
import { createComponent, internalState, type ComponentContext, type Inputs } from 'gesso-framework';

import type { Glyph } from './icons';
import { ICONS } from './icons';
import { Sheet } from './SheetContract';
import type { SheetEditing } from './SheetEditing';
import { saveSaid } from './whereItIs';

/**
 * The line along the very top: which workbook this is, and whether it
 * is safe — Phase 36.
 *
 * Until this the name was a word in the corner of the status bar, and
 * there was no way to say whether an edit had been kept. Every
 * spreadsheet people use daily puts both at the top: the name, which is
 * also where it is renamed, and a few words beside it that change from
 * *Saving…* to *Saved* — and here say *where*, because a workbook in
 * this application lives in the browser first and in a file only if
 * somebody saves one.
 *
 * One stop for the keyboard, the name, before the menu bar. The two
 * buttons beside it are File ▸ All workbooks and File ▸ Version
 * history, and are not stops for the reason the toolbar's buttons are
 * not: every stop here is one more press between the keyboard and the
 * sheet. The name commits on Enter or on leaving it, and Escape puts
 * back what it was.
 */
export interface DocumentBarProps {
  readonly editing: SheetEditing;
  readonly onHome: () => void;
  readonly onHistory: () => void;
  /** The name field, for File ▸ Rename to focus. */
  readonly nameRef?: (node: UiNode | null) => void;
}

export function DocumentBar(inputs: Inputs<DocumentBarProps>, ctx: ComponentContext) {
  const sheet = ctx.channel(Sheet);
  const edit = inputs.editing.value;
  const document = sheet.view.document;

  /** What the field holds while somebody types into it; null while nobody is. */
  let draft: string | null = null;
  const shown = internalState('');
  ctx.effect(
    document.pipe(
      map(current => current.name || 'Untitled'),
      distinctUntilChanged()
    ),
    name => {
      if (draft === null) {
        shown.value = name;
      }
    }
  );

  const commit = (): void => {
    const name = draft;
    draft = null;
    const current = document.value;
    if (name !== null && name.trim() !== '' && name.trim() !== current.name && current.id !== '') {
      sheet.send.renameDocument(current.id, name);
    } else {
      shown.value = current.name || 'Untitled';
    }
  };

  const onKey = (event: UiKeyboardEvent): void => {
    if (event.key === 'Enter') {
      event.preventDefault();
      event.stopPropagation();
      commit();
      edit.focusSheet();
    } else if (event.key === 'Escape') {
      event.preventDefault();
      event.stopPropagation();
      draft = null;
      shown.value = document.value.name || 'Untitled';
      edit.focusSheet();
    }
  };

  const said = document.pipe(map(saveSaid));

  return (
    <row width={percent(100)} y="center" gap={6} paddingLeft={6} paddingRight={10} paddingTop={4} paddingBottom={0}>
      {iconButton(ctx, ICONS.home, 'All workbooks', () => inputs.onHome.value())}
      <editabletext
        value={shown}
        ref={inputs.nameRef.value ?? undefined}
        width={shown.pipe(map(name => Math.min(360, Math.max(90, name.length * 8.5 + 18))))}
        fontSize={15}
        fontWeight="bold"
        color="text"
        textWrap="none"
        verticalAlign="middle"
        paddingLeft={6}
        paddingRight={6}
        paddingTop={2}
        paddingBottom={2}
        borderRadius={4}
        borderColor="transparent"
        borderWidth={1}
        role="textbox"
        label="Workbook name"
        modifiers={[tooltip(ctx, { text: 'Rename this workbook', placement: 'bottom' })]}
        onInput={(event: UiTextChangeEvent) => {
          draft = event.value;
          shown.value = event.value;
        }}
        onKeyDown={onKey}
        onBlur={() => {
          if (draft !== null) {
            commit();
          }
        }}
      />
      <text
        text={said.pipe(map(current => current.text))}
        fontSize={11}
        color={said.pipe(map(current => (current.tone === 'danger' ? 'danger' : 'textMuted')))}
        verticalAlign="middle"
        textWrap="none"
        selectable={false}
        role="status"
        label="Where this workbook is saved"
        live="polite"
      />
      <box flex={1} minWidth={0} />
      {iconButton(ctx, ICONS.history, 'Version history', () => inputs.onHistory.value())}
    </row>
  );
}

function iconButton(ctx: ComponentContext, glyph: Glyph, label: string, onClick: () => void) {
  return (
    <button
      focusable={false}
      onClick={onClick}
      label={label}
      modifiers={[tooltip(ctx, { text: label, placement: 'bottom' })]}
      padding={4}
      borderRadius={6}
      backgroundColor="transparent"
      borderColor="transparent"
      borderWidth={1}
      cursor="pointer">
      {createComponent(Icon, {
        path: glyph.path,
        viewBox: glyph.viewBox,
        size: 18,
        color: 'controlForeground',
        style: glyph.style,
        strokeWidth: glyph.strokeWidth ?? 1.5,
        fillRule: glyph.fillRule ?? 'nonzero'
      })}
    </button>
  );
}
