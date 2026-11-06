/** @jsxImportSource react */
// Shape controls: the insert-shape gallery (previews drawn by the engine's own geometry), the
// fill menu (theme colors, gradients, a pattern, and a picture), and the custom gradient dialog.

import { useEffect, useRef, useState, type PointerEvent as ReactPointerEvent } from 'react';
import {
  dialogInitialFocus,
  SHAPE_GALLERY,
  CROP_SHAPE_GALLERY,
  DEFAULT_GRADIENT,
  LINE_GLYPH_STROKE,
  defaultFillChoice,
  gradientAddStop,
  gradientAngleFrom,
  gradientDialogStart,
  gradientHandleKey,
  gradientPosAt,
  gradientRemoveStop,
  gradientSetStop,
  lineGlyphParts,
  lineGlyphTransform,
  lineSlotEnd,
  shapePreview,
  translateKeyOrText,
  type ChromeControl,
  type ChromeSlotId,
  type ColorChoice,
  type GradientFill,
  type PatternFill,
} from '@pptx-editor.dev/core/editor';
import { PatternDialog } from './PatternDialog.tsx';
import {
  Icon,
  MenuItem,
  Popover,
  ToolButton,
  focusText,
  keepFocus,
  useAriaDisabled,
} from './chrome.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';
import { useSlot } from './slot.ts';
import { useTriggerAttributes } from './slot-override.ts';
import { ColorSwatches, Swatch, SwatchGroup, useThemeBases } from './ColorSwatches.tsx';
import { colorMenuSelection, gradientCss, presetGradient } from '@pptx-editor.dev/core/editor';

/** The props of `ShapeIcon`. @public */
export interface ShapeIconProps {
  /** The preset geometry name (`a:prstGeom/@prst`), such as `rect`. */
  preset: string;
}

/** A preset's preview, drawn from its spec geometry. @public */
export function ShapeIcon({ preset }: ShapeIconProps) {
  const preview = shapePreview(preset);
  if (!preview) return null;
  return (
    <svg viewBox={preview.viewBox} className="pptx-shape-icon" aria-hidden="true">
      {preview.paths.map((p, i) => (
        <path
          key={i}
          d={p.d}
          className={p.fill ? 'pptx-shape-icon__fill' : undefined}
          fill={p.fill ? undefined : 'none'}
          stroke={p.stroke ? 'currentColor' : 'none'}
          vectorEffect="non-scaling-stroke"
        />
      ))}
    </svg>
  );
}

/** The props of `ShapeGrid`. @public */
export interface ShapeGridProps {
  /** The preset geometry names to show, in order. */
  presets: readonly string[];
  /** Called with the preset the person picked. */
  onPick: (preset: string) => void;
}

/**
 * One gallery category's presets as a grid of previews drawn from their spec geometry. Each
 * preview is a `menuitem`; the arrow keys move by grid rows (`data-columns`). @public
 */
export function ShapeGrid({ presets, onPick }: ShapeGridProps) {
  const { t } = useTranslation();
  return (
    <div className="pptx-shape-gallery__grid" data-columns={8}>
      {presets.map((preset) => {
        const name = translateKeyOrText(t, `shapes.${preset}`);
        return (
          <button
            key={preset}
            type="button"
            role="menuitem"
            tabIndex={-1}
            className="pptx-shape-gallery__item"
            aria-label={name}
            data-pptx-tooltip={name}
            onClick={() => onPick(preset)}
          >
            <ShapeIcon preset={preset} />
          </button>
        );
      })}
    </div>
  );
}

/** The props of `ShapeGalleryControl`. @public */
export interface ShapeGalleryControlProps {
  /** The chrome slot the control runs. */
  slot: ChromeSlotId;
  /** The registry entry of the control, from `CHROME_GROUPS`. */
  control: ChromeControl;
}

/** The insert-shape gallery: every category, every preset as its own preview. @public */
export function ShapeGalleryControl({ slot, control }: ShapeGalleryControlProps) {
  const { cmd, label, reason, tip } = useSlot(slot, control);
  const off = useAriaDisabled(!cmd.isEnabled, reason);
  const { t } = useTranslation();
  const anchor = useRef<HTMLButtonElement>(null);
  const [open, setOpen] = useState(false);
  const trigger = useTriggerAttributes(slot, 'pptx-toolbar__button pptx-toolbar__button--wide', {
    disabled: !cmd.isEnabled,
  });
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
        <Icon paths={control.paths} />
        <Icon name="arrow_drop_down" className="pptx-icon pptx-toolbar__caret" />
      </button>
      {off.note}
      {open ? (
        <Popover anchor={anchor.current} onClose={() => setOpen(false)}>
          <div className="pptx-shape-gallery">
            {/* A picture is cropped to shapes with an inside: lines are left out (core). */}
            {(slot === 'image.cropToShape' ? CROP_SHAPE_GALLERY : SHAPE_GALLERY).map((category) => (
              <section
                key={category.id}
                className="pptx-shape-gallery__section"
                role="group"
                aria-label={t(category.labelKey)}
              >
                <h3 className="pptx-shape-gallery__title" aria-hidden="true">
                  {t(category.labelKey)}
                </h3>
                <ShapeGrid
                  presets={category.presets}
                  onPick={(preset) => {
                    cmd.execute(preset);
                    setOpen(false);
                  }}
                />
              </section>
            ))}
          </div>
        </Popover>
      ) : null}
    </>
  );
}

/** The props of `FillControl`. @public */
export interface FillControlProps {
  /** The chrome slot the control runs. */
  slot: ChromeSlotId;
  /** The registry entry of the control, from `CHROME_GROUPS`. */
  control: ChromeControl;
}

/** The picture types a picture fill accepts. */
const PICTURE_TYPES = ['image/png', 'image/jpeg', 'image/gif'];

/**
 * Fill color: one menu with the theme and standard colors, Transparent, gradient presets from
 * the theme colors, a custom gradient dialog, a pattern dialog, and a picture. @public
 */
export function FillControl({ slot, control }: FillControlProps) {
  const { cmd, label, reason, tip } = useSlot(slot, control);
  const off = useAriaDisabled(!cmd.isEnabled, reason);
  const editor = usePptxEditor();
  const { t } = useTranslation();
  const current = useEditorState((s) => s.geometry?.gradient);
  const currentPattern = useEditorState((s) => s.geometry?.pattern);
  const [pattern, setPattern] = useState<PatternFill | null | false>(false);
  const file = useRef<HTMLInputElement>(null);
  const onFile = async (f: File | undefined): Promise<void> => {
    if (!f || !PICTURE_TYPES.includes(f.type)) return;
    let bytes: Uint8Array;
    try {
      bytes = new Uint8Array(await f.arrayBuffer());
    } catch {
      return; // An unreadable file fills nothing.
    }
    editor?.exec({ type: 'setShapePicture', bytes, mime: f.type });
    focusText(editor, anchor.current);
  };
  // The solid fill of the selection, marked in the menu.
  const selected = useEditorState((s) => colorMenuSelection(slot, s));
  const bases = useThemeBases();
  const anchor = useRef<HTMLButtonElement>(null);
  const [open, setOpen] = useState(false);
  const [picked, setPicked] = useState<{ choice: ColorChoice; color: string } | null>(null);
  // Before a pick, the button applies the theme's Accent 1, the default fill of a new shape.
  const accent1 = bases[4]?.color;
  const last: { choice: ColorChoice; color: string } = picked ?? defaultFillChoice(accent1);
  const [dialog, setDialog] = useState<GradientFill | null>(null);
  const applyGradient = (g: GradientFill): void => {
    editor?.exec({ type: 'setShapeGradient', gradient: g });
  };
  return (
    <span className="pptx-toolbar__group">
      <ToolButton
        slot={slot}
        label={label}
        paths={control.paths}
        disabled={!cmd.isEnabled}
        disabledReason={reason}
        onClick={() => cmd.execute(last.choice)}
      >
        <span className="pptx-toolbar__swatch" style={{ background: last.color }} />
      </ToolButton>
      <button
        ref={anchor}
        type="button"
        className="pptx-toolbar__button pptx-toolbar__step"
        aria-label={label}
        aria-haspopup="menu"
        aria-expanded={open}
        {...tip}
        {...off.attrs}
        onMouseDown={keepFocus}
        onClick={() => setOpen((o) => !o)}
      >
        <Icon name="arrow_drop_down" className="pptx-icon pptx-toolbar__caret" />
      </button>
      {off.note}
      {open ? (
        <Popover anchor={anchor.current} label={label} onClose={() => setOpen(false)}>
          <ColorSwatches
            selected={selected}
            onPick={(choice, color) => {
              setPicked({ choice, color });
              cmd.execute(choice);
              focusText(editor, anchor.current);
              setOpen(false);
            }}
          />
          <MenuItem
            onSelect={() => (cmd.execute(null), focusText(editor, anchor.current), setOpen(false))}
          >
            {t('toolbar.transparent')}
          </MenuItem>
          <div className="pptx-menu__sep" />
          {bases.length ? (
            <SwatchGroup label={t('toolbar.gradients')}>
              {bases.map((b) => {
                const g = presetGradient(b.color, 0.35, 0.35);
                return (
                  <Swatch
                    key={b.name}
                    color={b.color}
                    background={gradientCss(g)}
                    name={t('toolbar.gradientPreset', { color: b.name })}
                    menu
                    onPick={() => {
                      applyGradient(g);
                      focusText(editor, anchor.current);
                      setOpen(false);
                    }}
                  />
                );
              })}
            </SwatchGroup>
          ) : null}
          <MenuItem
            onSelect={() => {
              setDialog(current ?? presetGradient(last.color, 0.35, 0.35));
              setOpen(false);
            }}
          >
            <Icon name="add" />
            {t('toolbar.gradientCustom')}
          </MenuItem>
          <div className="pptx-menu__sep" />
          <MenuItem
            onSelect={() => {
              setPattern(currentPattern ?? null);
              setOpen(false);
            }}
          >
            {t('toolbar.pattern')}
          </MenuItem>
          <MenuItem
            onSelect={() => {
              file.current?.click();
              setOpen(false);
            }}
          >
            {t('toolbar.pictureFill')}
          </MenuItem>
        </Popover>
      ) : null}
      <input
        ref={file}
        type="file"
        accept={PICTURE_TYPES.join(',')}
        hidden
        onChange={(e) => {
          void onFile(e.target.files?.[0]);
          e.target.value = '';
        }}
      />
      <PatternDialog
        open={pattern !== false}
        initial={pattern || null}
        base={accent1}
        onClose={() => setPattern(false)}
        onApply={(p) => {
          editor?.exec({ type: 'setShapePattern', pattern: p });
          setPattern(false);
        }}
      />
      <GradientDialog
        open={dialog !== null}
        initial={dialog}
        onClose={() => setDialog(null)}
        onApply={(g) => {
          applyGradient(g);
          setDialog(null);
        }}
      />
    </span>
  );
}

/** Props for `PptxEditor.GradientDialog`. @public */
export interface GradientDialogProps {
  /** Whether the dialog shows. */
  readonly open: boolean;
  /** The gradient the dialog starts from. */
  readonly initial: GradientFill | null;
  /** Called when the dialog closes, applied or not. */
  readonly onClose: () => void;
  /** Called with the gradient the person applied. */
  readonly onApply: (gradient: GradientFill) => void;
}

/** The custom gradient dialog: type, angle, stops, rotation, and a preview. @public */
export function GradientDialog({ open, initial, onClose, onApply }: GradientDialogProps) {
  const { t } = useTranslation();
  const ref = useRef<HTMLDialogElement>(null);
  const bar = useRef<HTMLDivElement>(null);
  const [g, setG] = useState<GradientFill>(initial ?? DEFAULT_GRADIENT);
  const [selected, setSelected] = useState(0);
  const dragging = useRef<number | null>(null);

  useEffect(() => {
    const dialog = ref.current;
    if (!dialog) return;
    if (open) {
      setG(gradientDialogStart(initial));
      setSelected(0);
      if (!dialog.open) dialog.showModal();
      // The first field takes focus, as in every dialog, not the Close button.
      dialogInitialFocus(dialog)?.focus({ preventScroll: true });
    } else if (dialog.open) dialog.close();
  }, [open, initial]);

  const setStop = (i: number, patch: Partial<{ pos: number; color: string }>): void =>
    setG((cur) => gradientSetStop(cur, i, patch));
  const posAt = (clientX: number): number =>
    gradientPosAt(bar.current?.getBoundingClientRect(), clientX);
  const onHandleDown = (i: number) => (e: ReactPointerEvent<HTMLButtonElement>) => {
    setSelected(i);
    dragging.current = i;
    e.currentTarget.setPointerCapture(e.pointerId);
  };
  const onHandleMove = (e: ReactPointerEvent<HTMLButtonElement>) => {
    if (dragging.current !== null) setStop(dragging.current, { pos: posAt(e.clientX) });
  };
  const add = (): void => {
    const next = gradientAddStop(g, selected);
    setG(next.gradient);
    setSelected(next.selected);
  };
  const remove = (): void => {
    const next = gradientRemoveStop(g, selected);
    setG(next.gradient);
    setSelected(next.selected);
  };
  const stop = g.stops[Math.min(selected, g.stops.length - 1)];
  const css = gradientCss(g);

  return (
    <dialog
      ref={ref}
      className="pptx-dialog pptx-gradient"
      data-pptx-dialog="gradient"
      aria-labelledby="pptx-gradient-title"
      onCancel={(e) => {
        e.preventDefault();
        onClose();
      }}
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      {open ? (
        <>
          <div className="pptx-dialog__header">
            <h2 id="pptx-gradient-title" className="pptx-dialog__title">
              {t('gradient.title')}
            </h2>
            <ToolButton label={t('gradient.close')} icon="close" onClick={onClose} />
          </div>
          <div className="pptx-dialog__body">
            <div className="pptx-gradient__top">
              <div className="pptx-gradient__fields">
                <div className="pptx-gradient__row">
                  <label className="pptx-field">
                    {t('gradient.type')}
                    <select
                      value={g.kind}
                      onChange={(e) =>
                        setG({ ...g, kind: e.target.value === 'radial' ? 'radial' : 'linear' })
                      }
                    >
                      <option value="linear">{t('gradient.linear')}</option>
                      <option value="radial">{t('gradient.radial')}</option>
                    </select>
                  </label>
                  <label className="pptx-field">
                    {t('gradient.angle')}
                    <span className="pptx-field__control">
                      <input
                        type="number"
                        min={0}
                        max={359}
                        step={1}
                        value={Math.round(g.angle)}
                        disabled={g.kind === 'radial'}
                        onChange={(e) => {
                          const angle = gradientAngleFrom(e.target.value);
                          if (angle !== null) setG({ ...g, angle });
                        }}
                      />
                      <span className="pptx-field__unit">°</span>
                    </span>
                  </label>
                </div>
                <div className="pptx-dialog__section">
                  <h3 className="pptx-dialog__label">{t('gradient.stops')}</h3>
                  <div className="pptx-gradient__row">
                    <button type="button" className="pptx-button" onClick={add}>
                      {t('gradient.add')}
                    </button>
                    <button
                      type="button"
                      className="pptx-button"
                      onClick={remove}
                      disabled={g.stops.length <= 2}
                    >
                      {t('gradient.remove')}
                    </button>
                    <input
                      type="color"
                      className="pptx-gradient__color"
                      aria-label={t('gradient.stopColor')}
                      value={stop.color.toLowerCase()}
                      onChange={(e) => setStop(selected, { color: e.target.value.toUpperCase() })}
                    />
                  </div>
                </div>
                <label className="pptx-radio">
                  <input
                    type="checkbox"
                    checked={g.rotWithShape}
                    onChange={(e) => setG({ ...g, rotWithShape: e.target.checked })}
                  />
                  {t('gradient.rotate')}
                </label>
              </div>
              <div className="pptx-dialog__section">
                <h3 className="pptx-dialog__label">{t('gradient.preview')}</h3>
                <div className="pptx-gradient__preview" style={{ background: css }} />
              </div>
            </div>
            <div
              ref={bar}
              className="pptx-gradient__bar"
              style={{ background: gradientCss({ ...g, kind: 'linear', angle: 0 }) }}
            >
              {g.stops.map((s, i) => (
                <button
                  key={i}
                  type="button"
                  className="pptx-gradient__handle"
                  aria-label={t('gradient.stopHandle', { percent: Math.round(s.pos * 100) })}
                  aria-pressed={i === selected}
                  style={{ left: `${s.pos * 100}%`, background: s.color }}
                  onPointerDown={onHandleDown(i)}
                  onPointerMove={onHandleMove}
                  onPointerUp={() => (dragging.current = null)}
                  onKeyDown={(e) => {
                    const pos = gradientHandleKey(e.key, e.shiftKey, s.pos);
                    if (pos !== null) setStop(i, { pos });
                  }}
                />
              ))}
            </div>
          </div>
          <div className="pptx-dialog__footer">
            <button type="button" className="pptx-button" onClick={onClose}>
              {t('gradient.cancel')}
            </button>
            <button
              type="button"
              className="pptx-button pptx-button--primary"
              onClick={() => onApply(g)}
            >
              {t('gradient.ok')}
            </button>
          </div>
        </>
      ) : null}
    </dialog>
  );
}

/** The props of `LineGlyph`. @public */
export interface LineGlyphProps {
  /** The line kind or line end to draw, as the line menus give it. */
  value: string;
  /** Draw the end at the start of the line (mirrored), for the line start menu. */
  start?: boolean;
}

/** Glyphs for line kinds and line ends, drawn at the right of a 48×24 box. @public */
export function LineGlyph({ value, start }: LineGlyphProps) {
  const solid = { fill: 'currentColor', stroke: 'none' } as const;
  return (
    <svg className="pptx-line-glyph" viewBox="0 0 48 24" aria-hidden="true">
      <g transform={lineGlyphTransform(start)}>
        {lineGlyphParts(value, start).map((p, i) =>
          p.kind === 'dot' ? (
            <circle key={i} cx={p.cx} cy={p.cy} r={p.r} {...solid} />
          ) : p.kind === 'solid' ? (
            <path key={i} d={p.d} {...solid} />
          ) : (
            <path
              key={i}
              d={p.d}
              {...LINE_GLYPH_STROKE}
              {...(p.join ? { strokeLinejoin: 'round' as const } : {})}
            />
          )
        )}
      </g>
    </svg>
  );
}

/** The props of `LineMenuControl`. @public */
export interface LineMenuControlProps {
  /** The chrome slot the control runs. */
  slot: ChromeSlotId;
  /** The registry entry of the control, from `CHROME_GROUPS`. */
  control: ChromeControl;
}

/** The line tool and the line start and end menus, with each option drawn. @public */
export function LineMenuControl({ slot, control }: LineMenuControlProps) {
  const { cmd, label, reason, tip } = useSlot(slot, control);
  const off = useAriaDisabled(!cmd.isEnabled, reason);
  const editor = usePptxEditor();
  const { t } = useTranslation();
  const anchor = useRef<HTMLButtonElement>(null);
  const [open, setOpen] = useState(false);
  const isLine = useEditorState((s) => s.geometry?.lineHead !== undefined);
  // Line ends draw on the side they belong to; the line tool shows its kinds.
  const start = lineSlotEnd(slot);
  const shown = String(cmd.value ?? (start === undefined ? 'line' : 'none'));
  const trigger = useTriggerAttributes(slot, 'pptx-toolbar__button pptx-toolbar__button--wide', {
    disabled: !cmd.isEnabled,
  });
  // Line ends are contextual, as in Google Slides: they show only while a line is selected.
  if (start !== undefined && !isLine) return null;
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
        <LineGlyph value={shown} {...(start !== undefined ? { start } : {})} />
        <Icon name="arrow_drop_down" className="pptx-icon pptx-toolbar__caret" />
      </button>
      {off.note}
      {open ? (
        <Popover anchor={anchor.current} onClose={() => setOpen(false)}>
          {control.options?.map((o) => (
            <MenuItem
              key={o.value}
              // Line ends are a choice with a current value; the line tool inserts a kind.
              {...(control.toggle === 'choice' ? { checked: o.value === cmd.value } : {})}
              onSelect={() => (
                cmd.execute(o.value),
                focusText(editor, anchor.current),
                setOpen(false)
              )}
            >
              <LineGlyph value={o.value} {...(start !== undefined ? { start } : {})} />
              {t(o.labelKey)}
            </MenuItem>
          ))}
        </Popover>
      ) : null}
    </>
  );
}
