/** @jsxImportSource react */
// The color field of the form dialogs (Font color, Underline color, the bullet color): a list of
// the inherited color, "no change", and the theme colors, and beside it a swatch button that
// shows the current color and opens the color grid (theme tints and shades, standard
// colors). The picks and the color OK sends live in core (`color-fields.ts`).

import { useEffect, useRef, useState, type CSSProperties } from 'react';
import {
  colorChoiceOf,
  colorPickLabel,
  colorPickOf,
  colorPicks,
  focusCurrentSwatch,
  renderedColorOf,
  type ColorPick,
} from '@pptx-editor.dev/core/editor';
import { ColorSwatches } from './ColorSwatches.tsx';
import { useEditorState, useTranslation } from './context.ts';
import { SelectField, dialogControl } from './dialog-form.tsx';

/** A color pick as the slide renders it, or null for the inherited color and no change. */
function useRenderedColor(
  pick: ColorPick,
  current: string | null,
  automatic: string | null
): string | null {
  const rows = useEditorState((s) => s.themeColors);
  return renderedColorOf(pick, rows, current, automatic);
}

/** A color field: the list of picks and the swatch grid beside it. */
export function ColorField(props: {
  readonly name: string;
  readonly label: string;
  readonly value: ColorPick;
  /** The value the field opened with: `keep` and `mixed` are offered only then. */
  readonly seed: ColorPick | undefined;
  /** The color the selection has now (`#RRGGBB`): the swatch of "Current color" shows it. */
  readonly current?: string | null;
  /** The automatic color (`#RRGGBB`), when the field knows it: the swatch of `auto` shows it. */
  readonly automatic?: string | null;
  readonly disabled?: boolean;
  readonly onChange: (pick: ColorPick) => void;
  readonly className?: string;
  readonly style?: CSSProperties;
  readonly [attribute: `data-${string}`]: unknown;
}) {
  const { t } = useTranslation();
  const [open, setOpen] = useState(false);
  const panel = useRef<HTMLDivElement>(null);
  const button = useRef<HTMLButtonElement>(null);
  const rendered = useRenderedColor(props.value, props.current ?? null, props.automatic ?? null);
  const labelOf = (p: ColorPick): string => colorPickLabel(t, p);
  // The grid closes on a pick, on Escape (which does not close the dialog), and on a press
  // outside it.
  useEffect(() => {
    if (!open) return;
    const away = (e: PointerEvent): void => {
      const target = e.target as Node;
      if (!panel.current?.contains(target) && !button.current?.contains(target)) setOpen(false);
    };
    document.addEventListener('pointerdown', away, true);
    focusCurrentSwatch(panel.current);
    return () => document.removeEventListener('pointerdown', away, true);
  }, [open]);
  const data = Object.fromEntries(Object.entries(props).filter(([k]) => k.startsWith('data-')));
  return (
    <div
      {...data}
      className={['pptx-color-field', props.className].filter(Boolean).join(' ')}
      style={props.style}
    >
      <SelectField
        name={props.name}
        label={props.label}
        disabled={props.disabled}
        value={props.value}
        options={colorPicks(props.seed, props.value).map((p) => [p, labelOf(p)] as const)}
        onChange={props.onChange}
      />
      <button
        ref={button}
        type="button"
        className="pptx-swatch pptx-color-field__button"
        aria-label={t('colors.moreColors', { field: props.label })}
        aria-haspopup="true"
        aria-expanded={open}
        disabled={props.disabled}
        data-color={rendered ?? undefined}
        style={rendered ? { background: rendered } : undefined}
        onClick={() => setOpen((o) => !o)}
      />
      {open ? (
        <div
          ref={panel}
          role="group"
          aria-label={t('colors.moreColors', { field: props.label })}
          className="pptx-color-field__panel"
          onKeyDown={(e) => {
            if (e.key !== 'Escape') return;
            e.preventDefault();
            e.stopPropagation();
            setOpen(false);
            button.current?.focus();
          }}
        >
          <ColorSwatches
            menu={false}
            selected={colorChoiceOf(props.value)}
            onPick={(choice) => {
              props.onChange(colorPickOf(choice));
              setOpen(false);
              button.current?.focus();
            }}
          />
        </div>
      ) : null}
    </div>
  );
}

dialogControl(ColorField);
