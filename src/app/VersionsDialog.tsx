import { map } from 'rxjs';

import { Column, Text } from 'gesso-core';
import { Dialog } from 'gesso-components';
import { type ComponentContext, type Inputs } from 'gesso-framework';

import type { SheetVersionsView } from './SheetContract';
import { describeVersion } from './SheetVersions';
import { whenSaid } from './whereItIs';

/**
 * File ▸ Version history: the open workbook as it was, to put back —
 * Phase 38.
 *
 * Newest first, each with when and why it was kept. Restoring keeps
 * what is there now as a version of its own before replacing it, so
 * the list never loses a line by being used, and the sentence under it
 * says so.
 */
export interface VersionsDialogProps {
  readonly open: boolean;
  readonly versions: SheetVersionsView;
  readonly now?: () => number;
  readonly onRestore: (at: number) => void;
  readonly onClose: () => void;
}

export function VersionsDialog(inputs: Inputs<VersionsDialogProps>, _ctx: ComponentContext) {
  const now = (): number => (inputs.now.value ?? Date.now)();

  const button = (key: string, label: string, onClick: () => void, name = label) => (
    <button
      key={key}
      onClick={onClick}
      label={name}
      paddingLeft={10}
      paddingRight={10}
      paddingTop={4}
      paddingBottom={4}
      borderRadius={6}
      backgroundColor="controlBackground"
      borderColor="controlBorder"
      borderWidth={1}
      cursor="pointer">
      <text text={label} fontSize={12} textWrap="none" color="controlForeground" selectable={false} />
    </button>
  );

  const list = inputs.versions.pipe(
    map(versions => {
      if (versions.entries.length === 0) {
        return [
          <text
            key="none"
            text="No versions yet. One is kept the first time this workbook changes after it is opened, every ten minutes while it is edited, and whenever it is saved to a file."
            fontSize={12}
            color="textMuted"
            textWrap="word"
          />
        ];
      }
      return versions.entries.map(version => {
        const when = whenSaid(version.at, now());
        return (
          <row key={`v-${version.at}`} gap={8} y="center" minWidth={0} padding={3}>
            <text text={when} width={150} fontSize={12} fontWeight="bold" color="text" textWrap="none" />
            <text text={describeVersion(version.reason)} flex={1} minWidth={0} fontSize={12} color="textMuted" textWrap="none" />
            {button(`restore-${version.at}`, 'Restore', () => inputs.onRestore.value(version.at), `Restore the version from ${when}`)}
          </row>
        );
      });
    })
  );

  return (
    <Dialog
      open={inputs.open}
      onClose={() => inputs.onClose.value()}
      title="Version history"
      width={520}
      content={Column(
        { gap: 10, minWidth: 0 },
        <column gap={2} minWidth={0} maxHeight={320} overflow="auto" role="list" label="Versions">
          {list}
        </column>,
        Text({
          text: 'Restoring keeps the workbook as it is now as a version first, so nothing is lost by trying one. Versions are kept in this browser, thirty at most.',
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
