import { Column, Row, ScrollView, Text } from 'gesso-core';
import { Dialog } from 'gesso-components';

import { type ComponentContext, type Inputs } from 'gesso-framework';

import { acceleratorLabel, COMMANDS } from './SheetCommands';
import { NAVIGATION } from './SheetKeys';

/**
 * Every dialog here carries a Close button, and it is not decoration.
 *
 * `Dialog` traps focus into itself as it opens, and
 * `UiFocusManager.settleScope` *blurs* when the scope it is settling
 * into has nothing focusable in it. A dialog of plain text therefore
 * takes the keyboard away from whatever had it and hands it to
 * nothing: Escape reaches neither the sheet nor the dialog, and the
 * thing cannot be dismissed without a mouse. Found by pressing
 * Escape in a browser; `Overlays.spec.ts` upstream never catches it
 * because every dialog in it is given a button.
 *
 * A button is the honest fix rather than a workaround — a dialog
 * whose only exits are Escape and a backdrop click was never good —
 * but the engine's behaviour is the reason it is *required*, and that
 * is worth writing down next to the button.
 */

/**
 * The sheet of shortcuts, generated rather than written.
 *
 * Every row here is read out of a table that something else already
 * uses: the commands from `COMMANDS`, which is what the menus draw
 * from, and the navigation keys from `NAVIGATION`, which sits beside
 * the switch in `SheetKeys` that answers them. A help page written by
 * hand is a help page that is wrong within a month, and the version
 * that is wrong is worse than none because somebody believes it.
 */
export interface ShortcutsProps {
  readonly open: boolean;
  readonly onClose: () => void;
}

export function Shortcuts(inputs: Inputs<ShortcutsProps>, _ctx: ComponentContext) {
  /**
   * One line: what it does on the left, what to press on the right.
   *
   * The keys never wrap and never shrink. Left to the flex layout,
   * `Ctrl+Home / Ctrl+End` broke across two lines and dragged the
   * label into three, which is a help page that is harder to read
   * than the thing it explains.
   */
  const line = (label: string, keys: string) =>
    Row(
      { gap: 24, y: 'center', paddingTop: 3, paddingBottom: 3 },
      Text({ text: label, flex: 1, minWidth: 0, fontSize: 12, color: 'text', selectable: false }),
      Text({
        text: keys,
        flexShrink: 0,
        textWrap: 'none',
        fontSize: 11,
        color: 'textMuted',
        selectable: false
      })
    );

  const heading = (text: string) =>
    Text({
      text,
      fontSize: 11,
      fontWeight: 600,
      color: 'textMuted',
      marginTop: 10,
      marginBottom: 2,
      selectable: false
    });

  // From the command table rather than from the menus, so that a
  // command deliberately kept out of the bar — the key that opens the
  // bar — is still advertised. A key that works and is documented
  // nowhere is a key nobody presses.
  const commandLines = Object.values(COMMANDS)
    .filter(command => command.accelerator !== undefined)
    .map(command => line(command.label, acceleratorLabel(command.accelerator!)));

  return (
    <Dialog
      open={inputs.open}
      onClose={() => inputs.onClose.value()}
      title="Keyboard shortcuts"
      width={520}
      content={Column(
        { gap: 12, minWidth: 0 },
        ScrollView(
          {
          // Capped and scrolling, rather than as tall as it likes.
          // It was as tall as it liked, and the list is generated —
          // so it grew past the bottom of the screen and went on
          // drawing itself over the sheet. Every phase after this one
          // adds commands to it.
            maxHeight: 400,
            minWidth: 0
          },
          Column(
            { gap: 0, minWidth: 0 },
            heading('Moving around'),
            ...NAVIGATION.map(entry => line(entry.label, entry.keys)),
            heading('Commands'),
            ...commandLines
          )
        ),
        Row({ x: 'end' }, closeButton(() => inputs.onClose.value()))
      )}
    />
  );
}

/**
 * What the Paste menu item opens, because Paste cannot be a menu
 * item that pastes.
 *
 * `ShellService` can put text *on* the clipboard — that is the path
 * Phase 5's copy takes — and has no matching read, because reading
 * the clipboard is gated on a gesture inside a document and the
 * render worker has no document. The three options were a menu item
 * that silently does nothing, no Paste in the menu at all, and this.
 * The first is the worst thing software can do and the second sends
 * people looking. Google Sheets shows the same dialog, for the same
 * reason.
 */
export function PasteHint(inputs: Inputs<ShortcutsProps>, _ctx: ComponentContext) {
  return (
    <Dialog
      open={inputs.open}
      onClose={() => inputs.onClose.value()}
      title="Paste with the keyboard"
      width={380}
      content={Column(
        { gap: 12, minWidth: 0 },
        Text({
          text: 'Your browser only hands the clipboard to the keyboard shortcut, never to a menu.',
          fontSize: 12,
          color: 'text',
          textWrap: 'word'
        }),
        Text({
          text: `Press ${acceleratorLabel(COMMANDS.paste.accelerator!)} to paste.`,
          fontSize: 12,
          fontWeight: 600,
          color: 'text'
        }),
        Row({ x: 'end' }, closeButton(() => inputs.onClose.value()))
      )}
    />
  );
}

/**
 * The button that makes the dialog focusable, as the note at the top
 * of this file explains.
 */
function closeButton(onClick: () => void) {
  return (
    <button
      onClick={onClick}
      label="Close"
      paddingLeft={12}
      paddingRight={12}
      paddingTop={5}
      paddingBottom={5}
      borderRadius={6}
      backgroundColor="controlBackground"
      borderColor="controlBorder"
      borderWidth={1}
      cursor="pointer">
      <text text="Close" fontSize={12} color="controlForeground" selectable={false} />
    </button>
  );
}
