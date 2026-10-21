import { Column, Row, Text } from 'gesso-core';
import { Dialog } from 'gesso-components';
import { type ComponentContext, type Inputs } from 'gesso-framework';

import { TEMPLATES, type TemplateInfo, type TemplateKind } from './Templates';

/**
 * File ▸ New: a blank workbook, one started from a template, or an
 * example to look around.
 *
 * A dialog rather than three menu items, because the choice is the
 * point of it: a blank is one click as it always was, and the
 * templates and examples are in front of somebody at the moment they
 * are starting something, which is the only moment they are wanted.
 *
 * Every card opens here, as New does, with Back to return to the
 * workbook that was open; it is kept as it was. The cards are the ids in `Templates.ts`;
 * what each one writes is the application worker's, in
 * `SheetTemplates.ts`.
 */
export interface NewDialogProps {
  readonly open: boolean;
  /** A template's id, or `undefined` for a blank. */
  readonly onPick: (template: string | undefined) => void;
  readonly onClose: () => void;
}

export function NewDialog(inputs: Inputs<NewDialogProps>, _ctx: ComponentContext) {
  const pick = (template: string | undefined): void => inputs.onPick.value(template);

  const card = (key: string, name: string, blurb: string, onClick: () => void) => (
    <button
      key={key}
      onClick={onClick}
      label={name}
      flex={1}
      minWidth={0}
      x="start"
      y="start"
      paddingLeft={12}
      paddingRight={12}
      paddingTop={9}
      paddingBottom={9}
      borderRadius={8}
      backgroundColor="controlBackground"
      borderColor="controlBorder"
      borderWidth={1}
      cursor="pointer">
      {Column(
        { gap: 3, flex: 1, minWidth: 0 },
        Text({ text: name, fontSize: 12, fontWeight: 600, color: 'controlForeground', selectable: false }),
        Text({ text: blurb, fontSize: 11, color: 'textMuted', textWrap: 'word', selectable: false })
      )}
    </button>
  );

  const heading = (text: string) =>
    Text({ text, fontSize: 11, fontWeight: 600, color: 'textMuted', marginTop: 6, selectable: false });

  /**
   * The cards of one kind, two to a row. Pairs rather than a wrapping
   * row so each card is half the width exactly, whatever its blurb,
   * and an odd one out keeps its half rather than stretching.
   */
  const section = (kind: TemplateKind): ReturnType<typeof Row>[] => {
    const cards = TEMPLATES.filter(template => template.kind === kind);
    const rows: ReturnType<typeof Row>[] = [];
    for (let at = 0; at < cards.length; at += 2) {
      const pair: (TemplateInfo | undefined)[] = [cards[at], cards[at + 1]];
      rows.push(
        Row(
          { gap: 10, minWidth: 0 },
          ...pair.map((template, index) =>
            template === undefined
              ? <box key={`gap-${index}`} flex={1} minWidth={0} />
              : card(template.id, template.name, template.blurb, () => pick(template.id))
          )
        )
      );
    }
    return rows;
  };

  return (
    <Dialog
      open={inputs.open}
      onClose={() => inputs.onClose.value()}
      title="New workbook"
      description="Back returns to the workbook you are in now."
      width={560}
      content={Column(
        { gap: 10, minWidth: 0 },
        Row({ minWidth: 0 }, card('blank', 'Blank workbook', 'An empty sheet, to start from nothing.', () => pick(undefined))),
        heading('Templates'),
        ...section('template'),
        heading('Examples'),
        ...section('example'),
        <row gap={8} x="end" marginTop={4}>
          <button
            onClick={() => inputs.onClose.value()}
            label="Cancel"
            paddingLeft={12}
            paddingRight={12}
            paddingTop={5}
            paddingBottom={5}
            borderRadius={6}
            backgroundColor="controlBackground"
            borderColor="controlBorder"
            borderWidth={1}
            cursor="pointer">
            <text text="Cancel" fontSize={12} color="controlForeground" selectable={false} />
          </button>
        </row>
      )}
    />
  );
}
