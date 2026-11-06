// The top bar's two menus: the Editing | Viewing pill and Export (PPTX, PNG, ZIP, SVG). The Vue
// twins are examples/vue/src/ModePill.vue and examples/vue/src/ExportMenu.vue. Each trigger names
// its menu (`aria-haspopup`, `aria-expanded`, `aria-controls`); the library's `Popover` moves
// focus into the menu on a click too, and Escape returns it to the trigger. A row's name is its
// label; its hint is its description.

import { useId, useRef, useState } from 'react';
import { Icon, MenuItem, Popover, keepFocus, useTranslation } from '@pptx-editor.dev/react';
import { PICTURE_WIDTH, type SlideExport } from '../../shared/demo-files.ts';

export function ModePill({
  mode,
  onMode,
}: {
  mode: 'edit' | 'view';
  onMode: (m: 'edit' | 'view') => void;
}) {
  const { t } = useTranslation();
  const anchor = useRef<HTMLButtonElement>(null);
  const [open, setOpen] = useState(false);
  const id = useId();
  const current = mode === 'edit' ? t('mode.editing') : t('mode.viewing');
  return (
    <>
      <button
        ref={anchor}
        type="button"
        className="demo-modebar__pill"
        aria-haspopup="menu"
        aria-expanded={open}
        aria-controls={open ? id : undefined}
        // The name says the current mode: "Mode: Editing".
        aria-label={t('mode.current', { mode: current })}
        data-pptx-tooltip={t('mode.label')}
        onMouseDown={keepFocus}
        onClick={() => setOpen((o) => !o)}
      >
        <Icon name={mode === 'edit' ? 'edit_note' : 'visibility'} className="" />
        <span className="demo-modebar__label">{current}</span>
        <Icon name="arrow_drop_down" className="demo-btn__chev" />
      </button>
      {open ? (
        <Popover
          anchor={anchor.current}
          id={id}
          label={t('mode.label')}
          onClose={() => setOpen(false)}
        >
          {(['edit', 'view'] as const).map((m) => (
            <MenuItem
              key={m}
              checked={mode === m}
              description={m === 'edit' ? t('mode.editingHint') : t('mode.viewingHint')}
              onSelect={() => (onMode(m), setOpen(false))}
            >
              <span className="demo-menu__text">
                <span>{m === 'edit' ? t('mode.editing') : t('mode.viewing')}</span>
                <span className="demo-menu__hint" aria-hidden="true">
                  {m === 'edit' ? t('mode.editingHint') : t('mode.viewingHint')}
                </span>
              </span>
            </MenuItem>
          ))}
        </Popover>
      ) : null}
    </>
  );
}

const SLIDE_EXPORTS = [
  ['pdf', 'PDF', 'menu.exportPdf', 'menu.exportPdfHint'],
  ['png', 'PNG', 'menu.exportSlidePng', 'menu.exportSlidePngHint'],
  ['zip', 'ZIP', 'menu.exportSlidesPng', 'menu.exportSlidesPngHint'],
  ['svg', 'SVG', 'menu.exportSlideSvg', 'menu.exportSlideSvgHint'],
] as const;

export function ExportMenu({
  onExport,
  onExportSlides,
}: {
  onExport: () => void;
  onExportSlides: (kind: SlideExport) => void;
}) {
  const { t } = useTranslation();
  const anchor = useRef<HTMLButtonElement>(null);
  const [open, setOpen] = useState(false);
  const id = useId();
  return (
    <div className="demo-menu">
      <button
        ref={anchor}
        type="button"
        className="demo-btn demo-btn--sm demo-btn--primary"
        aria-haspopup="menu"
        aria-expanded={open}
        aria-controls={open ? id : undefined}
        aria-label={t('menu.export')}
        data-pptx-tooltip={t('menu.export')}
        onMouseDown={keepFocus}
        onClick={() => setOpen((o) => !o)}
      >
        <Icon name="file_download" className="" />
        <span className="demo-btn__label" aria-hidden="true">
          {t('menu.export')}
        </span>
        <Icon name="expand_more" className="demo-btn__chev" />
      </button>
      {open ? (
        <Popover
          anchor={anchor.current}
          id={id}
          label={t('menu.export')}
          onClose={() => setOpen(false)}
          align="end"
        >
          <MenuItem
            description={t('menu.exportPptxHint')}
            onSelect={() => (onExport(), setOpen(false))}
          >
            <span className="demo-filetype" aria-hidden="true">
              PPTX
            </span>
            <span className="demo-menu__text">
              <span>{t('menu.exportPptx')}</span>
              <span className="demo-menu__hint" aria-hidden="true">
                {t('menu.exportPptxHint')}
              </span>
            </span>
          </MenuItem>
          {SLIDE_EXPORTS.map(([kind, type, label, hint]) => (
            <MenuItem
              key={kind}
              description={t(hint, { width: PICTURE_WIDTH })}
              onSelect={() => (onExportSlides(kind), setOpen(false))}
            >
              <span className="demo-filetype" aria-hidden="true">
                {type}
              </span>
              <span className="demo-menu__text">
                <span>{t(label)}</span>
                <span className="demo-menu__hint" aria-hidden="true">
                  {t(hint, { width: PICTURE_WIDTH })}
                </span>
              </span>
            </MenuItem>
          ))}
        </Popover>
      ) : null}
    </div>
  );
}
