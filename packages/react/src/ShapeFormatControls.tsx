/** @jsxImportSource react */
// The Shape Format controls whose look is more than an icon: the Shape Styles gallery, which
// draws each theme style with the slide's theme colors, and the Shadow, Reflection, Glow, and
// Soft Edges menus, which name their current choice. Both are slot renderers: the registry
// says what they dispatch and when they are enabled (`chrome-shapes.ts`).

import { Fragment, useCallback, useRef, useState, type ReactNode } from 'react';
import {
  GLOW_ACCENTS,
  SHAPE_STYLE_GALLERY,
  chromeOptionChecked,
  effectSeparatorBefore,
  glowGridRows,
  shapeStyleChoice,
  shapeStylePreview,
  type ChromeControl,
  type ChromeSlotId,
  type ShapeStyleChoice,
  type ShapeStyleValue,
} from '@pptx-editor.dev/core/editor';
import { Icon, Popover, keepFocus, useAriaDisabled } from './chrome.tsx';
import { MenuRow, MenuSeparator, MenuSubmenu } from './ContextMenu.tsx';
import { useEditorState, useTranslation } from './context.ts';
import { useSlot } from './slot.ts';
import { useTriggerAttributes } from './slot-override.ts';

const cellKey = (c: ShapeStyleChoice): string => `${c.row}.${c.color}`;

/** The cell the trigger shows while the selection has no gallery style: Colored Fill, Accent 1. */
const DEFAULT_CELL: ShapeStyleChoice = { row: 'coloredFill', color: 'accent1' };

/** One gallery cell: a small rounded box in the style's fill, outline, and text colors. */
function StyleSwatch({
  cell,
  className = 'pptx-style-gallery__swatch',
}: {
  cell: ShapeStyleChoice;
  className?: string;
}): ReactNode {
  const { t } = useTranslation();
  const themeColors = useEditorState((s) => s.themeColors);
  const p = shapeStylePreview(cell, themeColors);
  return (
    <svg className={className} viewBox="0 0 40 28" aria-hidden="true">
      {p.effect > 0 ? (
        <rect
          x="4"
          y="5"
          width="34"
          height="22"
          rx="3"
          className="pptx-style-gallery__shadow"
          opacity={0.12 * p.effect}
        />
      ) : null}
      <rect
        x="2"
        y="2"
        width="34"
        height="22"
        rx="3"
        fill={p.fill ?? 'none'}
        stroke={p.line ?? 'none'}
        strokeWidth={p.line ? 1.5 : 0}
      />
      <text x="19" y="17" textAnchor="middle" fontSize="9" fill={p.text}>
        {t('toolbar.shapeStyleSample')}
      </text>
    </svg>
  );
}

/** Props of `ShapeStyleGallery`. @public */
export interface ShapeStyleGalleryProps {
  /**
   * The cell to check: a cell, or the slot value (`row.color`, as
   * `useEditorValueCommand('shape.styles').value` gives it), or null for none.
   */
  readonly value?: ShapeStyleChoice | ShapeStyleValue | null;
  /** Called with the picked cell. */
  readonly onPick: (cell: ShapeStyleChoice) => void;
  /** A class name added to the gallery. */
  readonly className?: string;
}

/**
 * The Quick Styles gallery: six theme style rows of seven colors, drawn with the current
 * slide's theme. Put it in a popover or a menu. @public
 */
export function ShapeStyleGallery({ value, onPick, className }: ShapeStyleGalleryProps): ReactNode {
  const picked = shapeStyleChoice(value);
  const checked = picked ? cellKey(picked) : undefined;
  const { t } = useTranslation();
  return (
    <div
      className={['pptx-style-gallery', className].filter(Boolean).join(' ')}
      role="group"
      aria-label={t('toolbar.shapeStyles')}
      // A grid for the arrow keys: Up and Down move by a row of colors.
      data-columns={SHAPE_STYLE_GALLERY[0]?.cells.length}
    >
      {SHAPE_STYLE_GALLERY.map(({ row, cells }) => (
        <div key={row} className="pptx-style-gallery__row" role="presentation">
          {cells.map((cell) => {
            const name = t('toolbar.shapeStyleCell', {
              row: t(`toolbar.shapeStyleRows.${cell.row}`),
              color: t(`toolbar.shapeStyleColors.${cell.color}`),
            });
            return (
              <button
                key={cellKey(cell)}
                type="button"
                role="menuitemradio"
                aria-checked={checked === cellKey(cell)}
                tabIndex={-1}
                className="pptx-menu__item pptx-style-gallery__cell"
                aria-label={name}
                data-pptx-tooltip={name}
                data-style={cellKey(cell)}
                onClick={() => onPick(cell)}
              >
                <StyleSwatch cell={cell} />
              </button>
            );
          })}
        </div>
      ))}
    </div>
  );
}

/** The toolbar control of `shape.styles`: a trigger and the gallery in a popover. @internal */
export function ShapeStyleGalleryControl({
  slot,
  control,
}: {
  slot: ChromeSlotId;
  control: ChromeControl;
}): ReactNode {
  const { cmd, label, reason, tip } = useSlot(slot, control);
  // A disabled picker stays in the arrow order and says why.
  const off = useAriaDisabled(!cmd.isEnabled, reason);
  const anchor = useRef<HTMLButtonElement>(null);
  const [open, setOpen] = useState(false);
  const trigger = useTriggerAttributes(slot, 'pptx-toolbar__button pptx-toolbar__button--wide', {
    disabled: !cmd.isEnabled,
  });
  const current = shapeStyleChoice(cmd.value) ?? DEFAULT_CELL;
  return (
    <>
      <button
        ref={anchor}
        type="button"
        {...trigger}
        data-slot={slot}
        aria-label={label}
        aria-haspopup="menu"
        aria-expanded={open}
        {...tip}
        {...off.attrs}
        onMouseDown={keepFocus}
        onClick={() => setOpen((o) => !o)}
      >
        <StyleSwatch cell={current} className="pptx-style-gallery__trigger" />
        <Icon name="arrow_drop_down" className="pptx-icon pptx-toolbar__caret" />
      </button>
      {off.note}
      {open ? (
        <Popover anchor={anchor.current} onClose={() => setOpen(false)}>
          <ShapeStyleGallery
            value={shapeStyleChoice(cmd.value)}
            onPick={(cell) => {
              cmd.execute(cell);
              setOpen(false);
            }}
          />
        </Popover>
      ) : null}
    </>
  );
}

/** The Shape Effects trigger: a rounded box with an offset shadow, drawn, not an icon path. */
function EffectsGlyph(): ReactNode {
  return (
    <svg className="pptx-icon pptx-effects-glyph" viewBox="0 0 24 24" aria-hidden="true">
      <rect x="7" y="7" width="13" height="13" rx="2.5" className="pptx-effects-glyph__shadow" />
      <rect x="4" y="4" width="13" height="13" rx="2.5" className="pptx-effects-glyph__box" />
    </svg>
  );
}

type EffectOption = NonNullable<ChromeControl['options']>[number];

/**
 * Glow Variations: a grid with a row per size and a column per accent, each cell a box with a
 * glow in that accent. Up and Down move by a row (`data-columns`).
 */
function GlowGrid({
  slot,
  options,
  value,
  onPick,
}: {
  slot: ChromeSlotId;
  options: readonly EffectOption[];
  value: unknown;
  onPick: (value: string) => void;
}): ReactNode {
  const { t } = useTranslation();
  const themeColors = useEditorState((s) => s.themeColors);
  const swatch = (accent: string): string =>
    themeColors[0]?.find((s) => s.slot === accent)?.color ?? '#808080';
  return (
    <div
      className="pptx-style-gallery pptx-glow-grid"
      role="group"
      aria-label={t('toolbar.glowVariationsLabel')}
      data-columns={GLOW_ACCENTS.length}
    >
      {glowGridRows(options).map(({ size, cells }) => (
        <div key={size} className="pptx-style-gallery__row" role="presentation">
          {cells.map(({ option: o, accent }) => {
            const name = t(o.labelKey, o.labelParams);
            return (
              <button
                key={o.value}
                type="button"
                role="menuitemradio"
                aria-checked={chromeOptionChecked(value, o.value)}
                tabIndex={-1}
                className="pptx-menu__item pptx-style-gallery__cell"
                aria-label={name}
                data-pptx-tooltip={name}
                data-slot={slot}
                data-value={o.value}
                onClick={() => onPick(o.value)}
              >
                <svg className="pptx-style-gallery__swatch" viewBox="0 0 40 28" aria-hidden="true">
                  <rect
                    x="9"
                    y="6"
                    width="22"
                    height="16"
                    rx="2"
                    fill="none"
                    stroke={swatch(accent)}
                    strokeOpacity="0.4"
                    // The glow width follows the size: 18 pt draws twice as wide as 8 pt.
                    strokeWidth={Number(size) / 2}
                  />
                  <rect x="9" y="6" width="22" height="16" rx="2" className="pptx-glow-grid__box" />
                </svg>
              </button>
            );
          })}
        </div>
      ))}
    </div>
  );
}

/**
 * The Shape Effects menu (`shape.effects`): one submenu per effect (Shadow, Reflection, Glow,
 * Soft Edges, the option groups of the registry), with each effect's current choice checked.
 * @internal
 */
export function EffectMenuControl({
  slot,
  control,
}: {
  slot: ChromeSlotId;
  control: ChromeControl;
}): ReactNode {
  const { cmd, label, reason, tip } = useSlot(slot, control);
  // A disabled picker stays in the arrow order and says why.
  const off = useAriaDisabled(!cmd.isEnabled, reason);
  const { t } = useTranslation();
  const anchor = useRef<HTMLButtonElement>(null);
  const [open, setOpen] = useState(false);
  const trigger = useTriggerAttributes(slot, 'pptx-toolbar__button pptx-toolbar__button--wide', {
    disabled: !cmd.isEnabled,
  });
  const close = useCallback(() => setOpen(false), []);
  return (
    <>
      <button
        ref={anchor}
        type="button"
        {...trigger}
        data-slot={slot}
        aria-label={label}
        aria-haspopup="menu"
        aria-expanded={open}
        {...tip}
        {...off.attrs}
        onMouseDown={keepFocus}
        onClick={() => setOpen((o) => !o)}
      >
        <EffectsGlyph />
        <Icon name="arrow_drop_down" className="pptx-icon pptx-toolbar__caret" />
      </button>
      {off.note}
      {open ? (
        <Popover anchor={anchor.current} onClose={close}>
          {control.optionGroups?.map((group) => {
            const options = control.options?.filter((o) => o.group === group.id) ?? [];
            const pick = (value: string) => (cmd.execute(value), close());
            const row = (o: EffectOption) => (
              <MenuRow
                key={o.value}
                slot={slot}
                label={t(o.labelKey, o.labelParams)}
                checked={chromeOptionChecked(cmd.value, o.value)}
                radio
                onSelect={() => pick(o.value)}
              />
            );
            const [first, ...rest] = options;
            return (
              <MenuSubmenu key={group.id} label={t(group.labelKey)}>
                {group.id === 'glow' && first ? (
                  <>
                    {row(first)}
                    <GlowGrid slot={slot} options={rest} value={cmd.value} onPick={pick} />
                  </>
                ) : (
                  // Shadows: a line between the outer and the inner presets (the Outer and
                  // Inner sections).
                  options.map((o, i) => (
                    <Fragment key={o.value}>
                      {effectSeparatorBefore(group.id, options, i) ? <MenuSeparator /> : null}
                      {row(o)}
                    </Fragment>
                  ))
                )}
              </MenuSubmenu>
            );
          })}
        </Popover>
      ) : null}
    </>
  );
}
