/** @jsxImportSource react */
// The Animations tab of the side panel: the slide transition, and the effects of the slide in
// play order. A row selects its effect and its shape; the fields below the list change the
// selected effect. The menus, labels, and seconds fields come from core (`animation-pane.ts`),
// so the Vue twin (`packages/vue/src/AnimationsPane.tsx`) offers the same choices. The editor
// does not play effects: the tab says so.

import { useEffect, useId, useState } from 'react';
import {
  ANIMATION_PRESET_GROUPS,
  ANIMATION_TRIGGERS,
  TRANSITION_KINDS,
  animationEffectLabelKey,
  animationListParts,
  parseSeconds,
  secondsText,
  transitionKindLabelKey,
  type AnimationEffect,
  type AnimationPreset,
  type AnimationTrigger,
  type TransitionKind,
} from '@pptx-editor.dev/core/editor';
import { ToolButton } from './chrome.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';

/** A seconds field that commits on Enter or blur and shows the document value otherwise. */
function SecondsField({
  label,
  ms,
  disabled,
  onCommit,
}: {
  label: string;
  ms: number;
  disabled: boolean;
  onCommit: (ms: number) => void;
}) {
  const { t } = useTranslation();
  const [draft, setDraft] = useState<string | null>(null);
  const shown = secondsText(ms);
  const commit = (): void => {
    const value = draft === null ? null : parseSeconds(draft);
    if (value !== null && draft !== shown) onCommit(value);
    setDraft(null);
  };
  return (
    <label className="pptx-field">
      {label}
      <span className="pptx-field__control">
        <input
          inputMode="decimal"
          aria-label={label}
          value={draft ?? shown}
          disabled={disabled}
          data-draft={draft !== null}
          onChange={(e) => setDraft(e.target.value)}
          onBlur={commit}
          onKeyDown={(e) => {
            if (e.key === 'Enter') commit();
            if (e.key === 'Escape') setDraft(null);
          }}
        />
        <span className="pptx-field__unit">{t('panel.animations.secondsUnit')}</span>
      </span>
    </label>
  );
}

/** The slide transition: kind, duration, and how the slide advances. */
function TransitionSection({ disabled }: { disabled: boolean }) {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const transition = useEditorState((s) => s.animations?.transition);
  const slideCount = useEditorState((s) => s.slides.length);
  const kindId = useId();
  if (!transition) return null;
  const set = (change: {
    kind?: TransitionKind | 'none';
    durationMs?: number;
    advanceOnClick?: boolean;
    advanceAfterMs?: number | null;
  }) => editor?.exec({ type: 'setTransition', ...change });
  return (
    <div className="pptx-panel__section">
      <h3 className="pptx-panel__title">{t('panel.animations.transition')}</h3>
      <label className="pptx-field" htmlFor={kindId}>
        {t('panel.animations.transitionKind')}
        <select
          id={kindId}
          className="pptx-animations__select"
          value={transition.kind}
          disabled={disabled}
          onChange={(e) => set({ kind: e.target.value as TransitionKind | 'none' })}
        >
          {transition.kind === 'other' ? (
            <option value="other" disabled>
              {t(transitionKindLabelKey('other'))}
            </option>
          ) : null}
          {TRANSITION_KINDS.map((kind) => (
            <option key={kind} value={kind}>
              {t(transitionKindLabelKey(kind))}
            </option>
          ))}
        </select>
      </label>
      <div className="pptx-panel__grid">
        <SecondsField
          label={t('panel.animations.duration')}
          ms={transition.durationMs}
          disabled={disabled || transition.kind === 'none'}
          onCommit={(ms) => set({ durationMs: Math.max(10, ms) })}
        />
        <SecondsField
          label={t('panel.animations.advanceAfter')}
          ms={transition.advanceAfterMs ?? 0}
          disabled={disabled || transition.advanceAfterMs === null}
          onCommit={(ms) => set({ advanceAfterMs: ms })}
        />
      </div>
      <div className="pptx-radios">
        <label className="pptx-radio">
          <input
            type="checkbox"
            checked={transition.advanceOnClick}
            disabled={disabled}
            onChange={(e) => set({ advanceOnClick: e.currentTarget.checked })}
          />
          {t('panel.animations.advanceOnClick')}
        </label>
        <label className="pptx-radio">
          <input
            type="checkbox"
            checked={transition.advanceAfterMs !== null}
            disabled={disabled}
            onChange={(e) => set({ advanceAfterMs: e.currentTarget.checked ? 0 : null })}
          />
          {t('panel.animations.advanceAfter')}
        </label>
      </div>
      <button
        type="button"
        className="pptx-button"
        disabled={disabled || slideCount < 2}
        onClick={() =>
          editor?.exec({
            type: 'setTransition',
            kind: transition.kind === 'other' ? undefined : transition.kind,
            durationMs: transition.durationMs,
            advanceOnClick: transition.advanceOnClick,
            advanceAfterMs: transition.advanceAfterMs,
            indices: Array.from({ length: slideCount }, (_, i) => i),
          })
        }
      >
        {t('panel.animations.applyToAll')}
      </button>
    </div>
  );
}

/** One row of the effect list: its start, its name, and its shape. */
function EffectRow({
  effect,
  selected,
  onSelect,
}: {
  effect: AnimationEffect;
  selected: boolean;
  onSelect: () => void;
}) {
  const { t } = useTranslation();
  const start =
    effect.step === null
      ? t('panel.animations.onShapeClick', {
          shape: effect.triggerShapeName || t('panel.animations.unnamed'),
        })
      : effect.step === 0
        ? t('panel.animations.withSlide')
        : t('panel.animations.step', { step: effect.step });
  return (
    <li className="pptx-animations__item">
      <button
        type="button"
        className="pptx-animations__row"
        aria-pressed={selected}
        data-class={effect.effectClass}
        onClick={onSelect}
      >
        <span className="pptx-animations__step">{effect.step ?? '•'}</span>
        <span className="pptx-animations__name">
          {t(animationEffectLabelKey(effect))}
          <span className="pptx-animations__detail">
            {effect.shapeName || t('panel.animations.unnamed')}
            {effect.paragraphs
              ? ` · ${t('panel.animations.paragraphs', {
                  start: effect.paragraphs.start + 1,
                  end: effect.paragraphs.end + 1,
                })}`
              : ''}
            {` · ${start}`}
          </span>
        </span>
      </button>
    </li>
  );
}

/** The fields of the selected effect: start, delay, duration, order, and removal. */
function EffectFields({
  effect,
  index,
  count,
  disabled,
}: {
  effect: AnimationEffect;
  index: number;
  count: number;
  disabled: boolean;
}) {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const startId = useId();
  const off = disabled || !effect.editable;
  const timing = (change: { trigger?: AnimationTrigger; delayMs?: number; durationMs?: number }) =>
    editor?.exec({ type: 'setAnimationTiming', effectId: effect.id, ...change });
  return (
    <div className="pptx-panel__section pptx-animations__fields">
      <label className="pptx-field" htmlFor={startId}>
        {t('panel.animations.start')}
        <select
          id={startId}
          className="pptx-animations__select"
          value={effect.trigger}
          disabled={off}
          onChange={(e) => timing({ trigger: e.target.value as AnimationTrigger })}
        >
          {ANIMATION_TRIGGERS.map((trigger) => (
            <option key={trigger} value={trigger}>
              {t(`panel.animations.triggers.${trigger}`)}
            </option>
          ))}
        </select>
      </label>
      <div className="pptx-panel__grid">
        <SecondsField
          label={t('panel.animations.delay')}
          ms={effect.delayMs}
          disabled={off}
          onCommit={(ms) => timing({ delayMs: ms })}
        />
        <SecondsField
          label={t('panel.animations.duration')}
          ms={effect.durationMs ?? 0}
          disabled={off || effect.durationMs === null}
          onCommit={(ms) => timing({ durationMs: Math.max(10, ms) })}
        />
      </div>
      <div className="pptx-animations__actions">
        <ToolButton
          label={t('panel.animations.moveEarlier')}
          icon="keyboard_arrow_up"
          disabled={off || index === 0}
          onClick={() =>
            editor?.exec({ type: 'moveAnimation', effectId: effect.id, to: index - 1 })
          }
        />
        <ToolButton
          label={t('panel.animations.moveLater')}
          icon="keyboard_arrow_down"
          disabled={off || index >= count - 1}
          onClick={() =>
            editor?.exec({ type: 'moveAnimation', effectId: effect.id, to: index + 1 })
          }
        />
        <ToolButton
          label={t('panel.animations.remove')}
          icon="delete"
          disabled={off}
          onClick={() => editor?.exec({ type: 'removeAnimation', effectId: effect.id })}
        />
      </div>
    </div>
  );
}

/** The Animations tab: the transition, the Add effect menu, and the effects of the slide. */
export function AnimationsPane() {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const animations = useEditorState((s) => s.animations);
  const readOnly = useEditorState((s) => s.status !== 'ready' || s.mode === 'view');
  const hasSelection = useEditorState((s) => s.selection.kind !== 'none');
  const [preset, setPreset] = useState<AnimationPreset>('fade');
  const [chosen, setChosen] = useState<string | null>(null);
  const presetId = useId();
  const listId = useId();
  // The canvas shows the order of the effects while the tab shows.
  useEffect(() => {
    if (!editor) return;
    editor.exec({ type: 'showAnimationOrder', show: true });
    return () => void editor.exec({ type: 'showAnimationOrder', show: false });
  }, [editor]);
  const effects = animations?.effects ?? [];
  const { main, other } = animationListParts(effects);
  const current = effects.find((e) => e.id === chosen) ?? null;
  const select = (effect: AnimationEffect): void => {
    setChosen(effect.id);
    if (effect.nodeId) editor?.setSelection({ kind: 'shapes', nodeIds: [effect.nodeId] });
  };
  const add = (): void => {
    const before = new Set(effects.map((e) => e.id));
    if (!editor?.exec({ type: 'addAnimation', preset }).ok) return;
    const added = editor.snapshot().animations?.effects.find((e) => !before.has(e.id));
    if (added) setChosen(added.id);
  };
  const editable = !readOnly && animations?.editable !== false;
  return (
    <div className="pptx-animations">
      <TransitionSection disabled={readOnly} />
      <div className="pptx-panel__section">
        <h3 className="pptx-panel__title" id={listId}>
          {t('panel.animations.effects')}
        </h3>
        <div className="pptx-animations__add">
          <label className="pptx-field" htmlFor={presetId}>
            {t('panel.animations.addEffect')}
            <select
              id={presetId}
              className="pptx-animations__select"
              value={preset}
              disabled={!editable}
              onChange={(e) => setPreset(e.target.value as AnimationPreset)}
            >
              {ANIMATION_PRESET_GROUPS.map((group) => (
                <optgroup
                  key={group.effectClass}
                  label={t(`panel.animations.classes.${group.effectClass}`)}
                >
                  {group.presets.map((p) => (
                    <option key={p} value={p}>
                      {t(`panel.animations.presets.${p}`)}
                    </option>
                  ))}
                </optgroup>
              ))}
            </select>
          </label>
          <button
            type="button"
            className="pptx-button"
            disabled={!editable || !hasSelection}
            onClick={add}
          >
            {t('panel.animations.add')}
          </button>
        </div>
        {animations && !animations.editable && effects.length ? (
          <p className="pptx-animations__note">{t('panel.animations.readOnly')}</p>
        ) : null}
        {effects.length === 0 ? (
          <div className="pptx-empty">
            <p className="pptx-empty__title">{t('panel.animations.emptyTitle')}</p>
            <p className="pptx-empty__body">{t('panel.animations.emptyBody')}</p>
          </div>
        ) : (
          <ol className="pptx-animations__list" aria-labelledby={listId}>
            {[...main, ...other].map((effect) => (
              <EffectRow
                key={`${effect.sequence}:${effect.id}`}
                effect={effect}
                selected={effect.id === current?.id}
                onSelect={() => select(effect)}
              />
            ))}
          </ol>
        )}
      </div>
      {current ? (
        <EffectFields
          effect={current}
          index={main.indexOf(current)}
          count={main.length}
          disabled={readOnly}
        />
      ) : null}
      <p className="pptx-animations__note">{t('panel.animations.noPlayback')}</p>
    </div>
  );
}
