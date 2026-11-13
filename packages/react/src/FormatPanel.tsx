/** @jsxImportSource react */
import { useEffect, useLayoutEffect, useRef, useState, type ReactNode } from 'react';
import {
  chromeControlForSlot,
  formatSectionsOf,
  formatControlMixed,
  formatTargetKey,
  formatSelectionCategory,
  formatSectionOpen,
  setFormatSectionOpen,
  listenFormatSectionExpansion,
  type FormatSectionId,
  type FormatSectionDefinition,
} from '@pptx-editor.dev/core/editor';
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';
import { useEditorDialog } from './popups.tsx';
import { ToolbarSlot } from './Toolbar.tsx';
import { FormatLineWidth } from './FormatLineWidth.tsx';
import { FormatFillTransparency } from './FormatFillTransparency.tsx';
import { FormatDetails } from './FormatDetails.tsx';
import { FormatGeometry } from './FormatGeometry.tsx';
import { ChartDataPane } from './ChartDataPane.tsx';
import { DiagramTextPane } from './DiagramTextPane.tsx';

/** Selection-aware format panel customization. @public */
export interface FormatPanelProps {
  /** Ordered section identifiers. Omitted identifiers are hidden. Defaults to all applicable sections. */
  readonly sections?: readonly FormatSectionId[];
  /** Replace an applicable section. A null value hides it. */
  readonly sectionContent?: Partial<Record<FormatSectionId, ReactNode>>;
  readonly className?: string;
  /** Content appended after built-in sections. */
  readonly children?: ReactNode;
}

/** Props of a reusable formatting section. @public */
export interface FormatSectionProps {
  readonly section: FormatSectionDefinition;
}

/** A built-in section, reusable in a custom inspector. @public */
export function FormatSection({ section }: FormatSectionProps) {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const background = useEditorDialog('background');
  const snapshot = useEditorState((s) => s);
  const [, refresh] = useState(0);
  useEffect(
    () => (editor ? listenFormatSectionExpansion(editor, () => refresh((n) => n + 1)) : undefined),
    [editor]
  );
  const category = formatSelectionCategory(snapshot);
  const expanded = formatSectionOpen(editor, category, section.id);
  return (
    <details className="pptx-format-disclosure" open={expanded}>
      <summary
        data-format-toggle={section.id}
        className="pptx-format-summary"
        onClick={(event) => {
          event.preventDefault();
          if (editor) setFormatSectionOpen(editor, category, section.id, !expanded);
        }}
      >
        {t(section.labelKey)}
      </summary>
      <div className="pptx-format-section-body">
        {section.id === 'geometry' ? <FormatGeometry /> : null}
        {section.id === 'fill' || section.id === 'effects' ? (
          <FormatDetails section={section.id} />
        ) : null}
        {section.id === 'geometry' ? null : section.id === 'chart' ? (
          <ChartDataPane />
        ) : section.id === 'diagram' ? (
          <DiagramTextPane />
        ) : (
          <div className="pptx-format-controls">
            {section.slots
              .filter((slot) => slot !== 'slide.background' || background.isAvailable)
              .flatMap((slot) => [
                <div className="pptx-format-control" data-format-slot={slot} key={slot}>
                  {slot !== 'shape.lineWidth' ? (
                    <span className="pptx-format-control__label">
                      {t(chromeControlForSlot(slot)!.labelKey)}
                    </span>
                  ) : null}
                  {formatControlMixed(editor, slot) ? (
                    <span className="pptx-format-mixed">{t('inspector.mixed')}</span>
                  ) : null}
                  {slot === 'shape.lineWidth' ? <FormatLineWidth /> : <ToolbarSlot slotId={slot} />}
                </div>,
                slot === 'shape.fill' ? <FormatFillTransparency key="fill.transparency" /> : null,
              ])}
          </div>
        )}
      </div>
    </details>
  );
}

/** Formats the current selection with shared commands and capability checks. @public */
export function FormatPanel({ sections, sectionContent, className, children }: FormatPanelProps) {
  const editor = usePptxEditor();
  const snapshot = useEditorState((s) => s);
  const { t } = useTranslation();
  const root = useRef<HTMLDivElement>(null);
  const category = formatSelectionCategory(snapshot);
  const previousCategory = useRef(category);
  useLayoutEffect(() => {
    if (previousCategory.current !== category) {
      const scroller = root.current?.closest<HTMLElement>('.pptx-panel__body');
      if (scroller) scroller.scrollTop = 0;
      previousCategory.current = category;
    }
  }, [category]);
  const target = `${editor?.session?.id ?? ''}:${formatTargetKey(snapshot)}`;
  return (
    <div ref={root} className={['pptx-format-panel', className].filter(Boolean).join(' ')}>
      {snapshot.selection.kind === 'shapes' && snapshot.selection.nodeIds.length > 1 ? (
        <p className="pptx-panel__title">
          {t('panel.format.multipleTitle', { count: snapshot.selection.nodeIds.length })}
        </p>
      ) : snapshot.geometry ? (
        <p className="pptx-panel__title">{snapshot.geometry.name}</p>
      ) : null}
      <div className="pptx-format-sections" key={editor?.session?.id}>
        {formatSectionsOf(snapshot, sections).map((section) => (
          <div key={section.id} data-format-section={section.id}>
            {sectionContent?.[section.id] !== undefined ? (
              sectionContent[section.id]
            ) : (
              <FormatSection
                key={section.id === 'diagram' ? section.id : target}
                section={section}
              />
            )}
          </div>
        ))}
      </div>
      {children}
    </div>
  );
}
