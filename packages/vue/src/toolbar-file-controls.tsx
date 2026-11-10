/** @jsxImportSource vue */
// The toolbar controls that open a popup or a file picker: the link control and the picture
// file controls (Insert Picture, Change Picture). The Vue twin of the React adapter's
// `toolbar-file-controls.tsx`.

import { defineComponent, inject, ref, type PropType } from 'vue';
import {
  captureImageTarget,
  imageFileTypes,
  insertImageFile,
  requestLinkPopup,
  translateKeyOrText,
  type ChromeControl,
  type ChromeSlotId,
} from '@pptx-editor.dev/core/editor';
import { ToolButton } from './chrome.tsx';
import { chromeActionHostKey, usePptxEditor, useTranslation } from './composables.ts';
import { useSlot } from './slot.ts';

const slotPartProps = {
  slot: { type: String as PropType<ChromeSlotId>, required: true as const },
  control: { type: Object as PropType<ChromeControl>, required: true as const },
} as const;

/** The link control: it opens the link popup (`popups.hyperlink`). @internal */
export const LinkControl = defineComponent({
  name: 'PptxLinkControl',
  props: slotPartProps,
  setup(props) {
    const { cmd, label, reason } = useSlot(
      () => props.slot,
      () => props.control
    );
    const editor = usePptxEditor();
    const { t } = useTranslation();
    const host = inject(chromeActionHostKey, null);
    const anchor = ref<HTMLElement | null>(null);
    return () => {
      // The button opens the popup, so a `popups.hyperlink` of `false` turns it off; the
      // command (`setLink`) keeps working.
      // Read the command state before `off`: a render that stops at `off` would not track it,
      // and the control would stay off after the reason goes away.
      const enabled = cmd.isEnabled.value;
      const why = reason.value;
      const off = host?.unavailable?.(props.slot) ?? null;
      return (
        <span ref={anchor} class="pptx-toolbar__group">
          <ToolButton
            slot={props.slot}
            label={label.value}
            paths={props.control.paths}
            disabled={!!off || !enabled}
            disabledReason={off ? translateKeyOrText(t, off) : why}
            pressed={typeof cmd.value.value === 'string' && cmd.value.value !== ''}
            onClick={() => requestLinkPopup(editor.value, anchor.value)}
          />
        </span>
      );
    };
  },
});

/** Insert Picture and Change Picture: a button over a hidden file input. @internal */
export const ImageControl = defineComponent({
  name: 'PptxImageControl',
  props: slotPartProps,
  setup(props) {
    const { cmd, label, reason } = useSlot(
      () => props.slot,
      () => props.control
    );
    const editor = usePptxEditor();
    const input = ref<HTMLInputElement | null>(null);
    let pickerTarget: (() => boolean) | undefined;
    return () => {
      // Insert > Picture adds a picture; Change Picture replaces the selected one's image.
      const target = props.control.state.kind === 'replace' ? 'replaceImage' : undefined;
      return [
        <ToolButton
          slot={props.slot}
          label={label.value}
          paths={props.control.paths}
          disabled={!cmd.isEnabled.value}
          disabledReason={reason.value}
          onClick={() => {
            pickerTarget = editor.value ? captureImageTarget(editor.value) : undefined;
            input.value?.click();
          }}
        />,
        <input
          ref={input}
          type="file"
          accept={imageFileTypes(target)}
          hidden
          onChange={(e: Event) => {
            const el = e.target as HTMLInputElement;
            const file = el.files?.[0];
            // Clear the input, so picking the same file again still fires `change`.
            el.value = '';
            const isCurrent = pickerTarget;
            pickerTarget = undefined;
            if (file && editor.value) void insertImageFile(editor.value, file, target, isCurrent);
          }}
        />,
      ];
    };
  },
});
