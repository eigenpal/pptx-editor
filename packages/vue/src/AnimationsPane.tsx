/** @jsxImportSource vue */
// The Animations tab of the side panel: the Vue twin of the React adapter's
// `AnimationsPane.tsx`, with the same markup, classes, and ARIA, on the same core helpers
// (`animation-pane.ts`). The editor does not play effects: the tab says so.

import { computed, defineComponent, onBeforeUnmount, ref, useId, watch, type PropType } from 'vue';
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
import { useEditorState, usePptxEditor, useTranslation } from './composables.ts';

/** A seconds field that commits on Enter or blur and shows the document value otherwise. */
const SecondsField = defineComponent({
  name: 'PptxAnimationSecondsField',
  props: {
    label: { type: String, required: true },
    ms: { type: Number, required: true },
    disabled: { type: Boolean, required: true },
    onCommit: { type: Function as PropType<(ms: number) => void>, required: true },
  },
  setup(props) {
    const { t } = useTranslation();
    const draft = ref<string | null>(null);
    const commit = (): void => {
      const d = draft.value;
      const value = d === null ? null : parseSeconds(d);
      if (value !== null && d !== secondsText(props.ms)) props.onCommit(value);
      draft.value = null;
    };
    return () => (
      <label class="pptx-field">
        {props.label}
        <span class="pptx-field__control">
          <input
            inputmode="decimal"
            aria-label={props.label}
            value={draft.value ?? secondsText(props.ms)}
            disabled={props.disabled}
            data-draft={draft.value !== null}
            onInput={(e: Event) => {
              draft.value = (e.target as HTMLInputElement).value;
            }}
            onBlur={commit}
            onKeydown={(e: KeyboardEvent) => {
              if (e.key === 'Enter') commit();
              if (e.key === 'Escape') draft.value = null;
            }}
          />
          <span class="pptx-field__unit">{t('panel.animations.secondsUnit')}</span>
        </span>
      </label>
    );
  },
});

/** The slide transition: kind, duration, and how the slide advances. */
const TransitionSection = defineComponent({
  name: 'PptxTransitionSection',
  props: { disabled: { type: Boolean, required: true } },
  setup(props) {
    const { t } = useTranslation();
    const editor = usePptxEditor();
    const transition = useEditorState((s) => s.animations?.transition);
    const slideCount = useEditorState((s) => s.slides.length);
    const kindId = useId();
    const set = (change: {
      kind?: TransitionKind | 'none';
      durationMs?: number;
      advanceOnClick?: boolean;
      advanceAfterMs?: number | null;
    }) => editor.value?.exec({ type: 'setTransition', ...change });
    return () => {
      const tr = transition.value;
      if (!tr) return null;
      const disabled = props.disabled;
      return (
        <div class="pptx-panel__section">
          <h3 class="pptx-panel__title">{t('panel.animations.transition')}</h3>
          <label class="pptx-field" for={kindId}>
            {t('panel.animations.transitionKind')}
            <select
              id={kindId}
              class="pptx-animations__select"
              value={tr.kind}
              disabled={disabled}
              onChange={(e: Event) =>
                set({ kind: (e.target as HTMLSelectElement).value as TransitionKind | 'none' })
              }
            >
              {tr.kind === 'other' ? (
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
          <div class="pptx-panel__grid">
            <SecondsField
              label={t('panel.animations.duration')}
              ms={tr.durationMs}
              disabled={disabled || tr.kind === 'none'}
              onCommit={(ms: number) => set({ durationMs: Math.max(10, ms) })}
            />
            <SecondsField
              label={t('panel.animations.advanceAfter')}
              ms={tr.advanceAfterMs ?? 0}
              disabled={disabled || tr.advanceAfterMs === null}
              onCommit={(ms: number) => set({ advanceAfterMs: ms })}
            />
          </div>
          <div class="pptx-radios">
            <label class="pptx-radio">
              <input
                type="checkbox"
                checked={tr.advanceOnClick}
                disabled={disabled}
                onChange={(e: Event) =>
                  set({ advanceOnClick: (e.currentTarget as HTMLInputElement).checked })
                }
              />
              {t('panel.animations.advanceOnClick')}
            </label>
            <label class="pptx-radio">
              <input
                type="checkbox"
                checked={tr.advanceAfterMs !== null}
                disabled={disabled}
                onChange={(e: Event) =>
                  set({
                    advanceAfterMs: (e.currentTarget as HTMLInputElement).checked ? 0 : null,
                  })
                }
              />
              {t('panel.animations.advanceAfter')}
            </label>
          </div>
          <button
            type="button"
            class="pptx-button"
            disabled={disabled || slideCount.value < 2}
            onClick={() =>
              editor.value?.exec({
                type: 'setTransition',
                kind: tr.kind === 'other' ? undefined : tr.kind,
                durationMs: tr.durationMs,
                advanceOnClick: tr.advanceOnClick,
                advanceAfterMs: tr.advanceAfterMs,
                indices: Array.from({ length: slideCount.value }, (_, i) => i),
              })
            }
          >
            {t('panel.animations.applyToAll')}
          </button>
        </div>
      );
    };
  },
});

/** One row of the effect list: its start, its name, and its shape. */
const EffectRow = defineComponent({
  name: 'PptxAnimationEffectRow',
  props: {
    effect: { type: Object as PropType<AnimationEffect>, required: true },
    selected: { type: Boolean, required: true },
    onSelect: { type: Function as PropType<() => void>, required: true },
  },
  setup(props) {
    const { t } = useTranslation();
    return () => {
      const effect = props.effect;
      const start =
        effect.step === null
          ? t('panel.animations.onShapeClick', {
              shape: effect.triggerShapeName || t('panel.animations.unnamed'),
            })
          : effect.step === 0
            ? t('panel.animations.withSlide')
            : t('panel.animations.step', { step: effect.step });
      return (
        <li class="pptx-animations__item">
          <button
            type="button"
            class="pptx-animations__row"
            aria-pressed={props.selected}
            data-class={effect.effectClass}
            onClick={() => props.onSelect()}
          >
            <span class="pptx-animations__step">{effect.step ?? '•'}</span>
            <span class="pptx-animations__name">
              {t(animationEffectLabelKey(effect))}
              <span class="pptx-animations__detail">
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
    };
  },
});

/** The fields of the selected effect: start, delay, duration, order, and removal. */
const EffectFields = defineComponent({
  name: 'PptxAnimationEffectFields',
  props: {
    effect: { type: Object as PropType<AnimationEffect>, required: true },
    index: { type: Number, required: true },
    count: { type: Number, required: true },
    disabled: { type: Boolean, required: true },
  },
  setup(props) {
    const { t } = useTranslation();
    const editor = usePptxEditor();
    const startId = useId();
    return () => {
      const effect = props.effect;
      const off = props.disabled || !effect.editable;
      const timing = (change: {
        trigger?: AnimationTrigger;
        delayMs?: number;
        durationMs?: number;
      }) => editor.value?.exec({ type: 'setAnimationTiming', effectId: effect.id, ...change });
      const move = (to: number) =>
        editor.value?.exec({ type: 'moveAnimation', effectId: effect.id, to });
      return (
        <div class="pptx-panel__section pptx-animations__fields">
          <label class="pptx-field" for={startId}>
            {t('panel.animations.start')}
            <select
              id={startId}
              class="pptx-animations__select"
              value={effect.trigger}
              disabled={off}
              onChange={(e: Event) =>
                timing({ trigger: (e.target as HTMLSelectElement).value as AnimationTrigger })
              }
            >
              {ANIMATION_TRIGGERS.map((trigger) => (
                <option key={trigger} value={trigger}>
                  {t(`panel.animations.triggers.${trigger}`)}
                </option>
              ))}
            </select>
          </label>
          <div class="pptx-panel__grid">
            <SecondsField
              label={t('panel.animations.delay')}
              ms={effect.delayMs}
              disabled={off}
              onCommit={(ms: number) => timing({ delayMs: ms })}
            />
            <SecondsField
              label={t('panel.animations.duration')}
              ms={effect.durationMs ?? 0}
              disabled={off || effect.durationMs === null}
              onCommit={(ms: number) => timing({ durationMs: Math.max(10, ms) })}
            />
          </div>
          <div class="pptx-animations__actions">
            <ToolButton
              label={t('panel.animations.moveEarlier')}
              icon="keyboard_arrow_up"
              disabled={off || props.index === 0}
              onClick={() => move(props.index - 1)}
            />
            <ToolButton
              label={t('panel.animations.moveLater')}
              icon="keyboard_arrow_down"
              disabled={off || props.index >= props.count - 1}
              onClick={() => move(props.index + 1)}
            />
            <ToolButton
              label={t('panel.animations.remove')}
              icon="delete"
              disabled={off}
              onClick={() => editor.value?.exec({ type: 'removeAnimation', effectId: effect.id })}
            />
          </div>
        </div>
      );
    };
  },
});

/** The Animations tab: the transition, the Add effect menu, and the effects of the slide. */
export const AnimationsPane = defineComponent({
  name: 'PptxAnimationsPane',
  setup() {
    const { t } = useTranslation();
    const editor = usePptxEditor();
    const animations = useEditorState((s) => s.animations);
    const readOnly = useEditorState((s) => s.status !== 'ready' || s.mode === 'view');
    const hasSelection = useEditorState((s) => s.selection.kind !== 'none');
    const preset = ref<AnimationPreset>('fade');
    const chosen = ref<string | null>(null);
    const presetId = useId();
    const listId = useId();
    // The canvas shows the order of the effects while the tab shows.
    watch(
      editor,
      (e, _old, onCleanup) => {
        if (!e) return;
        e.exec({ type: 'showAnimationOrder', show: true });
        onCleanup(() => void e.exec({ type: 'showAnimationOrder', show: false }));
      },
      { immediate: true }
    );
    onBeforeUnmount(() => void editor.value?.exec({ type: 'showAnimationOrder', show: false }));
    const effects = computed(() => animations.value?.effects ?? []);
    const add = (): void => {
      const e = editor.value;
      const before = new Set(effects.value.map((x) => x.id));
      if (!e?.exec({ type: 'addAnimation', preset: preset.value }).ok) return;
      const added = e.snapshot().animations?.effects.find((x) => !before.has(x.id));
      if (added) chosen.value = added.id;
    };
    const select = (effect: AnimationEffect): void => {
      chosen.value = effect.id;
      if (effect.nodeId) editor.value?.setSelection({ kind: 'shapes', nodeIds: [effect.nodeId] });
    };
    return () => {
      const list = effects.value;
      const { main, other } = animationListParts(list);
      const current = list.find((e) => e.id === chosen.value) ?? null;
      const editable = !readOnly.value && animations.value?.editable !== false;
      return (
        <div class="pptx-animations">
          <TransitionSection disabled={readOnly.value} />
          <div class="pptx-panel__section">
            <h3 class="pptx-panel__title" id={listId}>
              {t('panel.animations.effects')}
            </h3>
            <div class="pptx-animations__add">
              <label class="pptx-field" for={presetId}>
                {t('panel.animations.addEffect')}
                <select
                  id={presetId}
                  class="pptx-animations__select"
                  value={preset.value}
                  disabled={!editable}
                  onChange={(e: Event) => {
                    preset.value = (e.target as HTMLSelectElement).value as AnimationPreset;
                  }}
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
                class="pptx-button"
                disabled={!editable || !hasSelection.value}
                onClick={add}
              >
                {t('panel.animations.add')}
              </button>
            </div>
            {animations.value && !animations.value.editable && list.length ? (
              <p class="pptx-animations__note">{t('panel.animations.readOnly')}</p>
            ) : null}
            {list.length === 0 ? (
              <div class="pptx-empty">
                <p class="pptx-empty__title">{t('panel.animations.emptyTitle')}</p>
                <p class="pptx-empty__body">{t('panel.animations.emptyBody')}</p>
              </div>
            ) : (
              <ol class="pptx-animations__list" aria-labelledby={listId}>
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
              disabled={readOnly.value}
            />
          ) : null}
          <p class="pptx-animations__note">{t('panel.animations.noPlayback')}</p>
        </div>
      );
    };
  },
});
