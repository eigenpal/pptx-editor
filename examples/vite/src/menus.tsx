// The demo's menu bar, built on the library's `PptxEditor.Menu`: the built-in menus (File,
// Edit, View, Insert, Format, Arrange, Slide Show) come from the chrome registry, and the demo
// adds its own rows: opening and exporting files, slide structure, slide numbers,
// footers, dates, and Help.
// Every built-in row runs a registered slot, so its state and action are the toolbar's.

import { useEffect, useState, type ReactNode } from 'react';
import { PptxEditor, useEditorState, usePptxEditor, useTranslation } from '@pptx-editor.dev/react';
import { translateKeyOrText, type EditorCommand } from '@pptx-editor.dev/core/editor';
import { watchPdfExport, type SlideExport } from '../../shared/demo-files.ts';

/** A host row bound to a command: disabled with the engine's reason when `Editor.can` refuses it. */
function CommandRow({
  label,
  icon,
  command,
}: {
  label: string;
  icon: string;
  command: EditorCommand;
}): ReactNode {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  // A string, not the result object, so the subscription sees an unchanged answer as equal.
  const reason = useEditorState(() => {
    if (!editor) return 'disabled.noEditor';
    const can = editor.can(command);
    return can.ok ? null : can.reason;
  });
  return (
    <PptxEditor.Menu.Row
      label={label}
      icon={icon}
      disabled={!!reason}
      reason={reason ? translateKeyOrText(t, reason) : null}
      onSelect={() => editor?.exec(command)}
    />
  );
}

/**
 * Insert > Slide numbers: on or off for every slide, through each layout's slide-number
 * placeholder (`p:ph type="sldNum"`). The check shows whether the current slide has its number.
 */
function SlideNumbersRow(): ReactNode {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const on = useEditorState((s) => s.slideNumbers);
  const reason = useEditorState((s) => {
    if (!editor) return 'disabled.noEditor';
    const can = editor.can({ type: 'setSlideNumbers', show: !s.slideNumbers });
    return can.ok ? null : can.reason;
  });
  return (
    <PptxEditor.Menu.Row
      label={t('menu.slideNumbers')}
      icon="format_list_numbered"
      checked={on}
      disabled={!!reason}
      reason={reason ? translateKeyOrText(t, reason) : null}
      onSelect={() => editor?.exec({ type: 'setSlideNumbers', show: !on })}
    />
  );
}

/**
 * Insert > Footer: on or off for every slide, through each layout's footer placeholder
 * (`p:ph type="ftr"`). The check shows whether any slide has a footer, so the row never writes
 * over existing footers. A new footer says "Footer"; a host passes its own text.
 */
function FooterRow(): ReactNode {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const on = useEditorState((s) => (s.footer ?? null) !== null);
  const reason = useEditorState((s) => {
    if (!editor) return 'disabled.noEditor';
    const text = (s.footer ?? null) !== null ? null : t('menu.footer');
    const can = editor.can({ type: 'setFooter', text });
    return can.ok ? null : can.reason;
  });
  return (
    <PptxEditor.Menu.Row
      label={t('menu.footer')}
      icon="call_to_action"
      checked={on}
      disabled={!!reason}
      reason={reason ? translateKeyOrText(t, reason) : null}
      onSelect={() => editor?.exec({ type: 'setFooter', text: on ? null : t('menu.footer') })}
    />
  );
}

/**
 * Insert > Date and time: on or off for every slide, through each layout's date placeholder
 * (`p:ph type="dt"`). The date is a field that shows the date when the file opens.
 */
function DateTimeRow(): ReactNode {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const on = useEditorState((s) => s.dateTime === true);
  const reason = useEditorState((s) => {
    if (!editor) return 'disabled.noEditor';
    const can = editor.can({ type: 'setDateTime', show: s.dateTime !== true });
    return can.ok ? null : can.reason;
  });
  return (
    <PptxEditor.Menu.Row
      label={t('menu.dateTime')}
      icon="calendar_today"
      checked={on}
      disabled={!!reason}
      reason={reason ? translateKeyOrText(t, reason) : null}
      onSelect={() => editor?.exec({ type: 'setDateTime', show: !on })}
    />
  );
}

/** Insert > Duplicate slide: the current slide. */
function DuplicateSlideRow(): ReactNode {
  const { t } = useTranslation();
  const slide = useEditorState((s) => s.currentSlide);
  return (
    <CommandRow
      label={t('rail.duplicateSlide')}
      icon="content_copy"
      command={{ type: 'duplicateSlide', index: slide }}
    />
  );
}

export function MenuBar({
  onRename,
  onOpen,
  onNew,
  onExample,
  onExport,
  onExportSlides,
  onSelectionPane,
}: {
  onRename: () => void;
  onOpen: () => void;
  onSelectionPane: () => void;
  onNew: () => void;
  onExample: () => void;
  onExport: () => void;
  onExportSlides: (kind: SlideExport) => void;
}) {
  const { t } = useTranslation();
  // The PDF row waits while the server makes a PDF, as the Export menu's row does.
  const [pdfBusy, setPdfBusy] = useState(false);
  useEffect(() => watchPdfExport(setPdfBusy), []);
  return (
    <PptxEditor.Menu className="demo-menubar">
      {/* File: the demo's own file rows around the built-in Page setup. */}
      <PptxEditor.Menu.File preset={false}>
        <PptxEditor.Menu.Row label={t('menu.open')} icon="folder_open" onSelect={onOpen} />
        <PptxEditor.Menu.Row label={t('menu.new')} icon="note_add" onSelect={onNew} />
        <PptxEditor.Menu.Row label={t('menu.openSample')} icon="slideshow" onSelect={onExample} />
        <PptxEditor.Menu.Row label={t('menu.renameDeck')} icon="text_fields" onSelect={onRename} />
        <PptxEditor.Menu.Separator />
        <PptxEditor.Menu.Item slotId="file.pageSetup" />
        <PptxEditor.Menu.Separator />
        <PptxEditor.Menu.Group label={t('menu.export')}>
          <PptxEditor.Menu.Row
            label={t('menu.exportPptx')}
            icon="file_download"
            onSelect={onExport}
          />
          <PptxEditor.Menu.Row
            label={t('menu.exportPdf')}
            icon="description"
            disabled={pdfBusy}
            reason={t('menu.exportPdfBusy')}
            onSelect={() => onExportSlides('pdf')}
          />
          <PptxEditor.Menu.Row
            label={t('menu.exportSlidePng')}
            icon="image"
            onSelect={() => onExportSlides('png')}
          />
          <PptxEditor.Menu.Row
            label={t('menu.exportSlidesPng')}
            icon="folder_zip"
            onSelect={() => onExportSlides('zip')}
          />
          <PptxEditor.Menu.Row
            label={t('menu.exportSlideSvg')}
            icon="photo_library"
            onSelect={() => onExportSlides('svg')}
          />
        </PptxEditor.Menu.Group>
      </PptxEditor.Menu.File>
      {/* The slide rows go in Insert (host rows follow the registry rows here); delete and move live in the slide rail's menu. No extra Slide
          menu, so the bar fits a laptop window. */}
      <PptxEditor.Menu.Insert>
        <PptxEditor.Menu.Separator />
        {/* Duplicate slide: host rows follow the registry rows here, so it comes after the
            insert rows. */}
        <DuplicateSlideRow />
        {/* Layout and Reset often sit on a ribbon; this demo has no ribbon. The
            galleries draw each layout with the engine. In a group, the `slide.add` row is a
            second row next to the built-in New slide, not a replacement of it. */}
        <PptxEditor.Menu.Group>
          <PptxEditor.Menu.Item
            slotId="slide.add"
            direct={false}
            labelKey="rail.chooseLayout"
            icon="dashboard"
          />
          <PptxEditor.Menu.Item slotId="slide.layout" />
          <PptxEditor.Menu.Item slotId="slide.reset" />
        </PptxEditor.Menu.Group>
        <SlideNumbersRow />
        <FooterRow />
        <DateTimeRow />
      </PptxEditor.Menu.Insert>
      {/* Arrange > Selection pane opens the side panel on its Selection tab. */}
      <PptxEditor.Menu.Arrange>
        <PptxEditor.Menu.Separator />
        <PptxEditor.Menu.Row
          label={t('toolbar.selectionPane')}
          icon="view_sidebar"
          onSelect={onSelectionPane}
        />
      </PptxEditor.Menu.Arrange>
      <PptxEditor.Menu.Entry id="help" label={t('menu.help')}>
        <PptxEditor.Menu.Row
          label={t('menu.about')}
          icon="info"
          href="https://github.com/eigenpal/pptx-editor"
        />
        <PptxEditor.Menu.Row
          label={t('menu.reportIssue')}
          icon="bug_report"
          href="https://github.com/eigenpal/pptx-editor/issues"
        />
      </PptxEditor.Menu.Entry>
    </PptxEditor.Menu>
  );
}
