/** @jsxImportSource react */
// The swatch sections of every color menu and the background dialog: the deck's theme colors
// (ten columns, each with its lighter and darker variants), the Standard Colors row, and a
// custom color (a hex field and the browser's color selector), as DOCX Editor's color picker.

import {
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  type KeyboardEvent,
  type ReactNode,
} from 'react';
import type { ColorChoice } from '@pptx-editor.dev/core/editor';
import { useEditorState, useTranslation } from './context.ts';
import {
  STANDARD_COLORS,
  customHexColor,
  customHexCurrent,
  customHexDraft,
  customHexSeed,
  swatchGridKeyDown,
  swatchName,
  swatchSelection,
  syncSwatchTabStops,
  themeBaseSwatches,
  themeSwatchChoice,
} from '@pptx-editor.dev/core/editor';

/** One named swatch. In a menu it is a `menuitem`; in a dialog, a button that shows its state. */
export function Swatch({
  color,
  name,
  menu,
  pressed,
  background,
  onPick,
}: {
  color: string;
  name: string;
  menu: boolean;
  pressed?: boolean;
  /** CSS background when it is not the color itself (a gradient preview). */
  background?: string;
  onPick: () => void;
}) {
  return (
    <button
      type="button"
      className="pptx-swatch"
      style={{ background: background ?? color }}
      aria-label={name}
      data-pptx-tooltip={name}
      // In a menu the current color is marked with `aria-current`; in a dialog it is pressed.
      {...(menu
        ? { role: 'menuitem', tabIndex: -1, ...(pressed ? { 'aria-current': true } : {}) }
        : { 'aria-pressed': !!pressed })}
      onClick={onPick}
    />
  );
}

/**
 * A labelled swatch section; `columns` lets the menu's arrow keys move by row. With `roving`
 * (a dialog), the section is one Tab stop, on the chosen swatch or the first, and the arrow
 * keys move by column and row, as a WAI-ARIA grid of buttons does.
 */
export function SwatchGroup({
  label,
  columns = 10,
  className,
  roving = false,
  children,
}: {
  label: string;
  columns?: number;
  className?: string;
  roving?: boolean;
  children: ReactNode;
}) {
  const grid = useRef<HTMLDivElement>(null);
  // After each render: the focused swatch keeps the stop, else the chosen one, else the first.
  useLayoutEffect(() => {
    if (roving) syncSwatchTabStops(grid.current);
  });
  const onKeyDown = (e: KeyboardEvent<HTMLDivElement>): void => {
    if (roving) swatchGridKeyDown(grid.current, e, columns);
  };
  return (
    <div role="group" aria-label={label} className="pptx-swatches__section">
      <p className="pptx-swatches__title" aria-hidden="true">
        {label}
      </p>
      <div
        ref={grid}
        className={`pptx-swatches${className ? ` ${className}` : ''}`}
        data-columns={columns}
        onKeyDown={onKeyDown}
      >
        {children}
      </div>
    </div>
  );
}

/**
 * The Custom Color section, after DOCX Editor's color picker: a hex field (`#` and six digits)
 * with Apply, and the browser's color selector, which applies when it closes. In a menu, the
 * fields join the arrow keys (`data-menu-field`) and Apply is a row.
 */
export function CustomColor({
  onPick,
  menu,
  selected,
}: {
  onPick: (color: string) => void;
  menu: boolean;
  selected?: ColorChoice | null | undefined;
}) {
  const { t } = useTranslation();
  const [draft, setDraft] = useState(() => customHexSeed(selected));
  const color = customHexColor(draft);
  // An empty field shows the current color: as its placeholder and in the selector.
  const rows = useEditorState((s) => s.themeColors);
  const current = customHexCurrent(selected, rows);
  const apply = (): void => {
    if (color) onPick(color);
  };
  // The DOM `change` event: the selector closed on a color, which applies it.
  const picker = useRef<HTMLInputElement>(null);
  const pick = useRef(onPick);
  pick.current = onPick;
  useEffect(() => {
    const el = picker.current;
    const done = (): void => {
      const picked = el ? customHexColor(el.value) : null;
      if (picked) pick.current(picked);
    };
    el?.addEventListener('change', done);
    return () => el?.removeEventListener('change', done);
  }, []);
  return (
    <div role="group" aria-label={t('colors.customColor')} className="pptx-swatches__section">
      <p className="pptx-swatches__title" aria-hidden="true">
        {t('colors.customColor')}
      </p>
      <div className="pptx-custom-color">
        <input
          type="color"
          className="pptx-custom-color__picker"
          aria-label={t('colors.customPicker')}
          data-menu-field=""
          tabIndex={menu ? -1 : undefined}
          ref={picker}
          value={(color ?? `#${current}`).toLowerCase()}
          // React's `onChange` is the `input` event: it follows the selector while it moves.
          onChange={(e) => setDraft(customHexDraft(e.currentTarget.value))}
        />
        <span className="pptx-custom-color__hash" aria-hidden="true">
          #
        </span>
        <input
          type="text"
          className="pptx-custom-color__hex"
          aria-label={t('colors.customHex')}
          data-menu-field=""
          tabIndex={menu ? -1 : undefined}
          placeholder={current}
          maxLength={6}
          spellCheck={false}
          autoComplete="off"
          value={draft}
          onChange={(e) => setDraft(customHexDraft(e.target.value))}
          onKeyDown={(e) => {
            if (e.key !== 'Enter') return;
            // The Enter that applies is this field's: a dialog does not apply on it too.
            e.preventDefault();
            e.stopPropagation();
            apply();
          }}
        />
        <button
          type="button"
          className="pptx-button pptx-custom-color__apply"
          {...(menu ? { role: 'menuitem', tabIndex: -1 } : {})}
          disabled={!color}
          onClick={apply}
        >
          {t('colors.customApply')}
        </button>
      </div>
    </div>
  );
}

/**
 * The theme colors and the standard colors. A theme swatch picks a theme color, which the
 * engine writes as `a:schemeClr` (so the color follows a theme change); a
 * standard swatch and a custom color pick their `#RRGGBB`. `onPick` also gets the color as
 * rendered, for the split button's preview. `selected` marks the current color.
 */
export function ColorSwatches({
  onPick,
  menu = true,
  selected,
}: {
  onPick: (choice: ColorChoice, color: string) => void;
  menu?: boolean;
  selected?: ColorChoice | null;
}) {
  const { t } = useTranslation();
  const rows = useEditorState((s) => s.themeColors);
  // A theme choice matches its slot and shift; a hex choice marks one swatch only.
  const isSelected = swatchSelection(selected, rows);
  return (
    <>
      {rows.length ? (
        <SwatchGroup
          label={t('colors.themeColors')}
          className="pptx-swatches--theme"
          roving={!menu}
        >
          {rows.flatMap((row, r) =>
            row.map((s, c) => (
              <Swatch
                key={`${r}-${c}`}
                color={s.color}
                name={swatchName(t, s)}
                menu={menu}
                pressed={isSelected(s.color, s.slot, s.shift)}
                onPick={() => onPick(themeSwatchChoice(s), s.color)}
              />
            ))
          )}
        </SwatchGroup>
      ) : null}
      <SwatchGroup label={t('colors.standardColors')} roving={!menu}>
        {STANDARD_COLORS.map((s) => (
          <Swatch
            key={s.name}
            color={s.color}
            name={t(`colors.standard.${s.name}`)}
            menu={menu}
            pressed={isSelected(s.color)}
            onPick={() => onPick(s.color, s.color)}
          />
        ))}
      </SwatchGroup>
      <CustomColor menu={menu} selected={selected} onPick={(color) => onPick(color, color)} />
    </>
  );
}

/** The ten theme colors (the grid's first row), for gradient presets. */
export function useThemeBases(): readonly { color: string; name: string }[] {
  const { t } = useTranslation();
  const first = useEditorState((s) => s.themeColors[0]);
  return themeBaseSwatches(t, first);
}
