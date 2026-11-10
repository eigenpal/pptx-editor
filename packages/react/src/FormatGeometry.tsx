/** @jsxImportSource react */
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
import { useEditorDialog } from './popups.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';
import { FormatNumberField } from './FormatNumberField.tsx';
const EMU_PER_INCH = 914_400;
const inches = (emu: number): string => String(Number((emu / EMU_PER_INCH).toFixed(6)));

/** Position, size, rotation, and text fitting of the selected element. */
export function FormatGeometry() {
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
  if (!geometry)
    return selected > 1 ? (
      <div className="pptx-empty">
        <p className="pptx-empty__title">{t('panel.format.multipleTitle', { count: selected })}</p>
        <p className="pptx-empty__body">{platformShortcuts(t('panel.format.multipleBody'))}</p>
      </div>
    ) : (
      <>
        <div className="pptx-panel__section">
          <h3 className="pptx-panel__title">{t(panelTitle)}</h3>
          {background.isAvailable ? (
            <button
              type="button"
              className="pptx-button"
              disabled={readOnly}
              onClick={(e) => background.open(e.currentTarget)}
            >
              {t('menu.background')}
            </button>
          ) : null}
        </div>
        <div className="pptx-empty">
          <p className="pptx-empty__title">{t('panel.format.emptyTitle')}</p>
          <p className="pptx-empty__body">{t(panelBody)}</p>
        </div>
      </>
    );
  const g = geometry;
  const target = editor ? `${editor.session?.id}:${formatTargetKey(editor.snapshot())}` : '';
  const run = (command: EditorCommand): string | null => {
    if (!editor || `${editor.session?.id}:${formatTargetKey(editor.snapshot())}` !== target)
      return t('disabled.invalidValue');
    const result = editor.exec(command);
    return result.ok ? null : translateKeyOrText(t, result.reason);
  };
  const move = (patch: Partial<Record<'x' | 'y' | 'cx' | 'cy' | 'rot', number>>): string | null => {
    if (!editor || `${editor.session?.id}:${formatTargetKey(editor.snapshot())}` !== target)
      return t('disabled.invalidValue');
    const current = editor.snapshot().geometry;
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
  const field = (key: 'x' | 'y' | 'cx' | 'cy', labelKey: TranslationKey) => (
    <FormatNumberField
      label={t(labelKey)}
      property={`geometry.${key === 'cx' ? 'width' : key === 'cy' ? 'height' : key}`}
      min={key === 'cx' || key === 'cy' ? 0 : undefined}
      value={inches(g[key])}
      unit={unit}
      disabled={readOnly}
      // A locked aspect ratio sizes both sides from one.
      onCommit={(v) =>
        move(key === 'cx' || key === 'cy' ? formatPanelSize(g, key, inch(v)) : { [key]: inch(v) })
      }
    />
  );
  return (
    <>
      <div className="pptx-panel__section">
        <h3 className="pptx-panel__title">{t('panel.format.element')}</h3>
        <span className="pptx-chip">{t(elementKindLabelKey(g, table, { chart, diagram }))}</span>
      </div>
      <div className="pptx-panel__section">
        <h3 className="pptx-panel__title">{t('panel.format.size')}</h3>
        <div className="pptx-panel__grid">
          {field('cx', 'panel.format.width')}
          {field('cy', 'panel.format.height')}
        </div>
        <label className="pptx-radio">
          <input
            type="checkbox"
            checked={g.lockAspectRatio === true}
            disabled={readOnly}
            onChange={(e) =>
              editor?.exec({ type: 'setLockAspectRatio', lock: e.currentTarget.checked })
            }
          />
          {t('imageProperties.lockAspect')}
        </label>
      </div>
      <div className="pptx-panel__section">
        <h3 className="pptx-panel__title">{t('panel.format.position')}</h3>
        <div className="pptx-panel__grid">
          {field('x', 'panel.format.x')}
          {field('y', 'panel.format.y')}
        </div>
      </div>
      {elementRotates(g, table) ? (
        <div className="pptx-panel__section">
          <h3 className="pptx-panel__title">{t('panel.format.rotation')}</h3>
          <div className="pptx-panel__grid">
            <FormatNumberField
              label={t('panel.format.angle')}
              property="geometry.rotation"
              min={0}
              max={360}
              value={String(Math.round(g.rot * 10) / 10)}
              unit="°"
              disabled={readOnly}
              onCommit={(v) => move({ rot: v })}
            />
          </div>
        </div>
      ) : null}
      {table ? (
        <div className="pptx-panel__section">
          <h3 className="pptx-panel__title">{t('panel.format.table')}</h3>
          <div className="pptx-panel__grid">
            <FormatNumberField
              label={t('panel.format.rowHeight')}
              property="table-rowHeight"
              min={0}
              value={inches(table.rowHeight)}
              unit={unit}
              disabled={readOnly}
              onCommit={(v) => run({ type: 'setRowHeight', height: inch(v) })}
            />
            <FormatNumberField
              label={t('panel.format.columnWidth')}
              property="table-columnWidth"
              min={0}
              value={inches(table.columnWidth)}
              unit={unit}
              disabled={readOnly}
              onCommit={(v) => run({ type: 'setColumnWidth', width: inch(v) })}
            />
          </div>
          <div className="pptx-radios" role="group" aria-label={t('panel.format.styleOptions')}>
            {TABLE_OPTION_ROWS.map(({ option, labelKey }) => (
              <label key={option} className="pptx-radio">
                <input
                  type="checkbox"
                  checked={table.options.includes(option)}
                  disabled={readOnly}
                  onChange={() => editor?.exec({ type: 'toggleTableStyleOption', option })}
                />
                {t(labelKey)}
              </label>
            ))}
          </div>
        </div>
      ) : null}
      {g.textFitting ? (
        <div className="pptx-panel__section">
          <h3 className="pptx-panel__title">{t('panel.format.textFitting')}</h3>
          <div className="pptx-radios" role="radiogroup" aria-label={t('panel.format.textFitting')}>
            {(['none', 'shrink', 'shape'] as const).map((mode) => (
              <label key={mode} className="pptx-radio">
                <input
                  type="radio"
                  name="pptx-text-fitting"
                  checked={g.textFitting === mode}
                  disabled={readOnly}
                  onChange={() => editor?.exec({ type: 'setAutofit', mode })}
                />
                {t(`panel.format.fitting.${mode}`)}
              </label>
            ))}
          </div>
        </div>
      ) : null}
    </>
  );
}
