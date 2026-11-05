// One chrome slot as a control sees it: the engine's command state, the registry's label, and
// the tooltip that names the engine's reason when the control is disabled.

import { tooltipAttributes, translateKeyOrText } from '@pptx-editor.dev/core/editor';
import type {
  ChromeControl,
  ChromeSlotId,
  ChromeTooltipAttributes,
} from '@pptx-editor.dev/core/editor';

import { useEditorCommand, useTranslation, type UseEditorCommandReturn } from './context.ts';

/** Props of every slot part: the slot id and its registry entry. */
export interface SlotPartProps {
  readonly slot: ChromeSlotId;
  readonly control: ChromeControl;
}

/**
 * A slot's command, its translated label, and its tooltip attributes (with the reason on its
 * own line when disabled).
 */
export function useSlot(
  slot: ChromeSlotId,
  control: ChromeControl
): {
  readonly cmd: UseEditorCommandReturn;
  readonly label: string;
  readonly reason: string | null;
  readonly tip: ChromeTooltipAttributes;
} {
  const cmd = useEditorCommand(slot);
  const { t } = useTranslation();
  const label = t(control.labelKey);
  const reason = cmd.disabledReason ? translateKeyOrText(t, cmd.disabledReason) : null;
  return { cmd, label, reason, tip: tooltipAttributes(label, !cmd.isEnabled, reason) };
}
