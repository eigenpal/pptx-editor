// One chrome slot as a control sees it: the engine's command state, the registry's label, and
// the tooltip that names the engine's reason when the control is disabled. The Vue twin of the
// React adapter's `slot.ts`.

import { computed, toValue, type ComputedRef, type MaybeRefOrGetter } from 'vue';
import {
  tooltipAttributes,
  translateKeyOrText,
  type ChromeControl,
  type ChromeSlotId,
  type ChromeTooltipAttributes,
} from '@pptx-editor.dev/core/editor';
import { useEditorCommand, useTranslation, type UseEditorCommandReturn } from './composables.ts';

/** Props of every slot part: the slot id and its registry entry. @internal */
export interface SlotPartProps {
  readonly slot: ChromeSlotId;
  readonly control: ChromeControl;
}

/** What `useSlot` returns. @internal */
export interface UseSlotReturn {
  readonly cmd: UseEditorCommandReturn;
  readonly label: ComputedRef<string>;
  readonly reason: ComputedRef<string | null>;
  readonly tip: ComputedRef<ChromeTooltipAttributes>;
}

/** A slot's command, its translated label, and its tooltip attributes (with the reason when disabled). @internal */
export function useSlot(
  slot: MaybeRefOrGetter<ChromeSlotId>,
  control: MaybeRefOrGetter<ChromeControl>
): UseSlotReturn {
  const cmd = useEditorCommand(slot);
  const { t } = useTranslation();
  const label = computed(() => t(toValue(control).labelKey));
  const reason = computed(() =>
    cmd.disabledReason.value ? translateKeyOrText(t, cmd.disabledReason.value) : null
  );
  const tip = computed(() => tooltipAttributes(label.value, !cmd.isEnabled.value, reason.value));
  return { cmd, label, reason, tip };
}
