import { combineLatest, map } from 'rxjs';

import { Column, Text, type UiKeyboardEvent, type UiNode, type UiTextChangeEvent } from 'gesso-core';
import { Dialog } from 'gesso-components';
import { FocusService, internalState, type ComponentContext, type Inputs } from 'gesso-framework';

/**
 * Writing a cell's note: Shift+F2, or Insert ▸ Note.
 *
 * A dialog rather than a box on the sheet, because a note is a few
 * lines and a box the size of a cell is the wrong shape to write them
 * in. Enter is a new line, as it is in any note; Ctrl+Enter saves, and
 * so does the button. Saving an empty note takes the note away, which
 * is also what Delete does, for somebody who looks for the word.
 */
export interface NoteDialogProps {
  readonly open: boolean;
  /** Which cell, as the title says it: `B4`. */
  readonly cell: string;
  /** What the note says now, or empty for a cell with none. */
  readonly text: string;
  readonly onSave: (text: string) => void;
  readonly onClose: () => void;
}

export function NoteDialog(inputs: Inputs<NoteDialogProps>, ctx: ComponentContext) {
  const focus = ctx.inject(FocusService);
  const draft = internalState('');
  // A fresh draft each time the dialog opens, from the note as it is.
  ctx.effect(combineLatest([inputs.open, inputs.text]), ([open, text]) => {
    if (open) {
      draft.value = text;
    }
  });

  const save = (): void => inputs.onSave.value(draft.value);

  const onKey = (event: UiKeyboardEvent): void => {
    if (event.key === 'Enter' && (event.modifiers.ctrl === true || event.modifiers.meta === true)) {
      event.preventDefault();
      event.stopPropagation();
      save();
    }
  };

  const button = (label: string, onClick: () => void) => (
    <button
      onClick={onClick}
      label={label}
      paddingLeft={12}
      paddingRight={12}
      paddingTop={5}
      paddingBottom={5}
      borderRadius={6}
      backgroundColor="controlBackground"
      borderColor="controlBorder"
      borderWidth={1}
      cursor="pointer">
      <text text={label} fontSize={12} color="controlForeground" selectable={false} />
    </button>
  );

  return (
    <Dialog
      open={inputs.open}
      onClose={() => inputs.onClose.value()}
      title="Note"
      width={360}
      content={Column(
        { gap: 12, minWidth: 0 },
        Text({ text: inputs.cell.pipe(map(cell => `On ${cell}`)), fontSize: 12, color: 'textMuted' }),
        <editabletext
          value={draft}
          multiline={true}
          ref={(node: UiNode | null) => {
            if (node !== null) {
              focus.focus(node);
            }
          }}
          minHeight={96}
          fontSize={12}
          color="text"
          textWrap="word"
          backgroundColor="background"
          borderColor="border"
          borderWidth={1}
          padding={6}
          role="textbox"
          label="Note"
          onInput={(event: UiTextChangeEvent) => (draft.value = event.value)}
          onKeyDown={onKey}
        />,
        <row gap={8} x="end">
          {inputs.text.pipe(map(text => (text === '' ? [] : [button('Delete note', () => inputs.onSave.value(''))])))}
          {button('Cancel', () => inputs.onClose.value())}
          {button('Save', save)}
        </row>
      )}
    />
  );
}
