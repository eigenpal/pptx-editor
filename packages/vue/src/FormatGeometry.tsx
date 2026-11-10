/** @jsxImportSource vue */
import { defineComponent } from 'vue';
import {
  TABLE_OPTION_ROWS,
  elementKindLabelKey,
  elementRotates,
  formatPanelSize,
  platformShortcuts,
  formatPanelKeys,
  formatTargetKey,
  translateKeyOrText,
  type EditorCommand,
} from '@pptx-editor.dev/core/editor';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import { useEditorDialog } from './popups.ts';
import { useEditorState, usePptxEditor, useTranslation } from './composables.ts';
import { FormatNumberField } from './FormatNumberField.tsx';
const EMU_PER_INCH = 914_400;
const inches = (emu: number): string => String(Number((emu / EMU_PER_INCH).toFixed(6)));

/** Position, size, rotation, and text fitting of the selected element. */
export const FormatGeometry = defineComponent({
  name: 'PptxFormatTab',
  setup() {
    const { t } = useTranslation();
    const editor = usePptxEditor();
    const geometry = useEditorState((s) => s.geometry);
    const table = useEditorState((s) => s.table);
    const chart = useEditorState((s) => s.chart ?? null);
    const diagram = useEditorState((s) => s.diagram ?? null);
    const readOnly = useEditorState((s) => s.mode === 'view');
    // A layout or a master in the slide master view names itself in the heading.
    const panelTitle = useEditorState((s) => formatPanelKeys(s).title);
    const panelBody = useEditorState((s) => formatPanelKeys(s).body);
    const selected = useEditorState((s) =>
      s.selection.kind === 'shapes' ? s.selection.nodeIds.length : 0
    );
    const background = useEditorDialog('background');
    return () => {
      const g = geometry.value;
      if (!g)
        return selected.value > 1 ? (
          <div class="pptx-empty">
            <p class="pptx-empty__title">
              {t('panel.format.multipleTitle', { count: selected.value })}
            </p>
            <p class="pptx-empty__body">{platformShortcuts(t('panel.format.multipleBody'))}</p>
          </div>
        ) : (
          [
            <div class="pptx-panel__section">
              <h3 class="pptx-panel__title">{t(panelTitle.value)}</h3>
              {background.isAvailable.value ? (
                <button
                  type="button"
                  class="pptx-button"
                  disabled={readOnly.value}
                  onClick={(e: MouseEvent) => background.open(e.currentTarget as HTMLElement)}
                >
                  {t('menu.background')}
                </button>
              ) : null}
            </div>,
            <div class="pptx-empty">
              <p class="pptx-empty__title">{t('panel.format.emptyTitle')}</p>
              <p class="pptx-empty__body">{t(panelBody.value)}</p>
            </div>,
          ]
        );
      const target = editor.value
        ? `${editor.value.session?.id}:${formatTargetKey(editor.value.snapshot())}`
        : '';
      const run = (command: EditorCommand): string | null => {
        if (
          !editor.value ||
          `${editor.value.session?.id}:${formatTargetKey(editor.value.snapshot())}` !== target
        )
          return t('disabled.invalidValue');
        const result = editor.value.exec(command);
        return result.ok ? null : translateKeyOrText(t, result.reason);
      };
      const move = (
        patch: Partial<Record<'x' | 'y' | 'cx' | 'cy' | 'rot', number>>
      ): string | null => {
        if (
          !editor.value ||
          `${editor.value.session?.id}:${formatTargetKey(editor.value.snapshot())}` !== target
        )
          return t('disabled.invalidValue');
        const current = editor.value.snapshot().geometry;
        if (!current || current.nodeId !== g.nodeId) return t('disabled.invalidValue');
        return run({
          type: 'setTransform',
          nodeId: g.nodeId,
          x: patch.x ?? current.x,
          y: patch.y ?? current.y,
          cx: patch.cx ?? current.cx,
          cy: patch.cy ?? current.cy,
          rot: patch.rot ?? current.rot,
        });
      };
      const inch = (v: number): number => Math.round(v * EMU_PER_INCH);
      const unit = t('panel.format.inchUnit');
      const tableSection = () => {
        const tb = table.value;
        if (!tb) return null;
        return (
          <div class="pptx-panel__section">
            <h3 class="pptx-panel__title">{t('panel.format.table')}</h3>
            <div class="pptx-panel__grid">
              <FormatNumberField
                label={t('panel.format.rowHeight')}
                property="table-rowHeight"
                min={0}
                value={inches(tb.rowHeight)}
                unit={unit}
                disabled={readOnly.value}
                onCommit={(v: number) => run({ type: 'setRowHeight', height: inch(v) })}
              />
              <FormatNumberField
                label={t('panel.format.columnWidth')}
                property="table-columnWidth"
                min={0}
                value={inches(tb.columnWidth)}
                unit={unit}
                disabled={readOnly.value}
                onCommit={(v: number) => run({ type: 'setColumnWidth', width: inch(v) })}
              />
            </div>
            <div class="pptx-radios" role="group" aria-label={t('panel.format.styleOptions')}>
              {TABLE_OPTION_ROWS.map(({ option, labelKey }) => (
                <label key={option} class="pptx-radio">
                  <input
                    type="checkbox"
                    checked={tb.options.includes(option)}
                    disabled={readOnly.value}
                    onChange={() => editor.value?.exec({ type: 'toggleTableStyleOption', option })}
                  />
                  {t(labelKey)}
                </label>
              ))}
            </div>
          </div>
        );
      };
      const field = (key: 'x' | 'y' | 'cx' | 'cy', labelKey: TranslationKey) => (
        <FormatNumberField
          label={t(labelKey)}
          property={`geometry.${key === 'cx' ? 'width' : key === 'cy' ? 'height' : key}`}
          min={key === 'cx' || key === 'cy' ? 0 : undefined}
          value={inches(g[key])}
          unit={unit}
          disabled={readOnly.value}
          // A locked aspect ratio sizes both sides from one.
          onCommit={(v: number) =>
            move(
              key === 'cx' || key === 'cy' ? formatPanelSize(g, key, inch(v)) : { [key]: inch(v) }
            )
          }
        />
      );
      return [
        <div class="pptx-panel__section">
          <h3 class="pptx-panel__title">{t('panel.format.element')}</h3>
          <span class="pptx-chip">
            {t(elementKindLabelKey(g, table.value, { chart: chart.value, diagram: diagram.value }))}
          </span>
        </div>,
        <div class="pptx-panel__section">
          <h3 class="pptx-panel__title">{t('panel.format.size')}</h3>
          <div class="pptx-panel__grid">
            {field('cx', 'panel.format.width')}
            {field('cy', 'panel.format.height')}
          </div>
          <label class="pptx-radio">
            <input
              type="checkbox"
              checked={g.lockAspectRatio === true}
              disabled={readOnly.value}
              onChange={(e: Event) =>
                editor.value?.exec({
                  type: 'setLockAspectRatio',
                  lock: (e.currentTarget as HTMLInputElement).checked,
                })
              }
            />
            {t('imageProperties.lockAspect')}
          </label>
        </div>,
        <div class="pptx-panel__section">
          <h3 class="pptx-panel__title">{t('panel.format.position')}</h3>
          <div class="pptx-panel__grid">
            {field('x', 'panel.format.x')}
            {field('y', 'panel.format.y')}
          </div>
        </div>,
        elementRotates(g, table.value) ? (
          <div class="pptx-panel__section">
            <h3 class="pptx-panel__title">{t('panel.format.rotation')}</h3>
            <div class="pptx-panel__grid">
              <FormatNumberField
                label={t('panel.format.angle')}
                property="geometry.rotation"
                min={0}
                max={360}
                value={String(Math.round(g.rot * 10) / 10)}
                unit="°"
                disabled={readOnly.value}
                onCommit={(v: number) => move({ rot: v })}
              />
            </div>
          </div>
        ) : null,
        tableSection(),
        g.textFitting ? (
          <div class="pptx-panel__section">
            <h3 class="pptx-panel__title">{t('panel.format.textFitting')}</h3>
            <div class="pptx-radios" role="radiogroup" aria-label={t('panel.format.textFitting')}>
              {(['none', 'shrink', 'shape'] as const).map((mode) => (
                <label key={mode} class="pptx-radio">
                  <input
                    type="radio"
                    name="pptx-text-fitting"
                    checked={g.textFitting === mode}
                    disabled={readOnly.value}
                    onChange={() => editor.value?.exec({ type: 'setAutofit', mode })}
                  />
                  {t(`panel.format.fitting.${mode}`)}
                </label>
              ))}
            </div>
          </div>
        ) : null,
      ];
    };
  },
});
