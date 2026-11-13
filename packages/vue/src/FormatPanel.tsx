/** @jsxImportSource vue */
import {
  defineComponent,
  ref,
  watch,
  type ExtractPublicPropTypes,
  type PropType,
  type SlotsType,
  type VNodeChild,
} from 'vue';
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
import { useEditorState, usePptxEditor, useTranslation } from './composables.ts';
import { useEditorDialog } from './popups.ts';
import { ToolbarSlot } from './toolbar.ts';
import { FormatLineWidth } from './FormatLineWidth.tsx';
import { FormatFillTransparency } from './FormatFillTransparency.tsx';
import { FormatDetails } from './FormatDetails.tsx';
import { FormatGeometry } from './FormatGeometry.tsx';
import { ChartDataPane } from './ChartDataPane.tsx';
import { DiagramTextPane } from './DiagramTextPane.tsx';

const formatSectionProps = {
  section: { type: Object as PropType<FormatSectionDefinition>, required: true as const },
};
/** Props of a reusable formatting section. @public */
export type FormatSectionProps = ExtractPublicPropTypes<typeof formatSectionProps>;

/** A built-in section, reusable in a custom inspector. @public */
export const FormatSection = defineComponent({
  name: 'FormatSection',
  props: formatSectionProps,
  setup(props) {
    const { t } = useTranslation();
    const editor = usePptxEditor();
    const background = useEditorDialog('background');
    const snapshot = useEditorState((s) => s);
    const expansion = ref(0);
    watch(
      editor,
      (instance, _old, onCleanup) => {
        if (instance)
          onCleanup(
            listenFormatSectionExpansion(instance, () => {
              expansion.value += 1;
            })
          );
      },
      { immediate: true }
    );
    return () => {
      void expansion.value;
      const section = props.section;
      const category = formatSelectionCategory(snapshot.value);
      const expanded = formatSectionOpen(editor.value, category, section.id);
      return (
        <details class="pptx-format-disclosure" open={expanded}>
          <summary
            data-format-toggle={section.id}
            class="pptx-format-summary"
            onClick={(event: MouseEvent) => {
              event.preventDefault();
              if (editor.value) setFormatSectionOpen(editor.value, category, section.id, !expanded);
            }}
          >
            {t(section.labelKey)}
          </summary>
          <div class="pptx-format-section-body">
            {section.id === 'geometry' ? <FormatGeometry /> : null}
            {section.id === 'fill' || section.id === 'effects' ? (
              <FormatDetails section={section.id} />
            ) : null}
            {section.id === 'geometry' ? null : section.id === 'chart' ? (
              <ChartDataPane />
            ) : section.id === 'diagram' ? (
              <DiagramTextPane />
            ) : (
              <div class="pptx-format-controls">
                {section.slots
                  .filter((slot) => slot !== 'slide.background' || background.isAvailable.value)
                  .flatMap((slot) => [
                    <div class="pptx-format-control" data-format-slot={slot} key={slot}>
                      {slot !== 'shape.lineWidth' ? (
                        <span class="pptx-format-control__label">
                          {t(chromeControlForSlot(slot)!.labelKey)}
                        </span>
                      ) : null}
                      {formatControlMixed(editor.value, slot) ? (
                        <span class="pptx-format-mixed">{t('inspector.mixed')}</span>
                      ) : null}
                      {slot === 'shape.lineWidth' ? (
                        <FormatLineWidth />
                      ) : (
                        <ToolbarSlot slotId={slot} />
                      )}
                    </div>,
                    slot === 'shape.fill' ? (
                      <FormatFillTransparency key="fill.transparency" />
                    ) : null,
                  ])}
              </div>
            )}
          </div>
        </details>
      );
    };
  },
});

const formatPanelProps = {
  sections: { type: Array as PropType<readonly FormatSectionId[]>, default: undefined },
  className: { type: String, default: undefined },
};
/** Props of the selection-aware format panel. @public */
export type FormatPanelProps = ExtractPublicPropTypes<typeof formatPanelProps>;

/** Selection-aware formatting. Named section slots replace their packaged sections. @public */
export const FormatPanel = defineComponent({
  name: 'FormatPanel',
  props: formatPanelProps,
  slots: Object as SlotsType<Partial<Record<FormatSectionId | 'default', () => VNodeChild>>>,
  setup(props, { slots }) {
    const editor = usePptxEditor();
    const snapshot = useEditorState((s) => s);
    const { t } = useTranslation();
    const root = ref<HTMLDivElement | null>(null);
    watch(
      () => formatSelectionCategory(snapshot.value),
      (category, previous) => {
        if (category !== previous) {
          const scroller = root.value?.closest<HTMLElement>('.pptx-panel__body');
          if (scroller) scroller.scrollTop = 0;
        }
      },
      { flush: 'post' }
    );
    return () => (
      <div ref={root} class={['pptx-format-panel', props.className]}>
        {snapshot.value.selection.kind === 'shapes' &&
        snapshot.value.selection.nodeIds.length > 1 ? (
          <p class="pptx-panel__title">
            {t('panel.format.multipleTitle', { count: snapshot.value.selection.nodeIds.length })}
          </p>
        ) : snapshot.value.geometry ? (
          <p class="pptx-panel__title">{snapshot.value.geometry.name}</p>
        ) : null}
        <div class="pptx-format-sections" key={editor.value?.session?.id}>
          {formatSectionsOf(snapshot.value, props.sections).map((section) => (
            <div key={section.id} data-format-section={section.id}>
              {slots[section.id] ? (
                slots[section.id]!()
              ) : (
                <FormatSection
                  key={section.id === 'diagram' ? section.id : formatTargetKey(snapshot.value)}
                  section={section}
                />
              )}
            </div>
          ))}
        </div>
        {slots.default?.()}
      </div>
    );
  },
});
