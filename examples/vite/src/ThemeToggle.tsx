// The light/dark toggle in the top bar, with the shape of the DOCX Editor demos' toggle. The
// Vue twin is examples/vue/src/ThemeToggle.vue; both use the `demo-theme-toggle` styles in
// examples/shared/shell.css.

import { radioGroupStep } from '@pptx-editor.dev/core/editor';
import { useTranslation } from '@pptx-editor.dev/react';

type Theme = 'light' | 'dark';
const MODES = ['light', 'dark'] as const;

const SUN =
  'M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4';
const MOON = 'M12 3a6 6 0 0 0 9 9 9 9 0 1 1-9-9Z';

export function ThemeToggle({ value, onChange }: { value: Theme; onChange: (t: Theme) => void }) {
  const { t } = useTranslation();
  return (
    <div
      className="demo-theme-toggle"
      role="radiogroup"
      aria-label={t('pageSetup.appearance')}
      // One Tab stop: the checked option; the arrow keys move the choice, as a radio group does.
      onKeyDown={(e) => {
        const next = radioGroupStep(MODES, value, e.key);
        if (!next) return;
        e.preventDefault();
        onChange(next);
        e.currentTarget.querySelector<HTMLElement>(`[data-value="${next}"]`)?.focus();
      }}
    >
      {MODES.map((mode) => (
        <button
          key={mode}
          type="button"
          role="radio"
          className="demo-theme-toggle__option"
          data-value={mode}
          aria-checked={value === mode}
          tabIndex={value === mode ? 0 : -1}
          aria-label={t(mode === 'light' ? 'pageSetup.light' : 'pageSetup.dark')}
          title={t(mode === 'light' ? 'pageSetup.light' : 'pageSetup.dark')}
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => onChange(mode)}
        >
          <svg viewBox="0 0 24 24" aria-hidden="true">
            {mode === 'light' ? <circle cx="12" cy="12" r="4" /> : null}
            <path d={mode === 'light' ? SUN : MOON} />
          </svg>
        </button>
      ))}
    </div>
  );
}
