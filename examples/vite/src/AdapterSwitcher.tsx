// The React | Vue switcher, as in the DOCX Editor demos. The preview deploy (`bun run
// build:preview`) serves the React demo at `/react/` and the Vue demo at `/vue/`; in development
// the two demos run on their own ports (5274 and 5275). The Vue twin is
// examples/vue/src/AdapterSwitcher.vue; both use the styles in examples/shared/shell.css.
//
// A navigation, not tabs: the other demo is a link that keeps the page's query (`?theme=dark`,
// `?chrome=ribbon`, `?deck=…`), and the current demo is marked `aria-current="page"` and is not
// a link, so the switcher is one Tab stop.

import { useTranslation } from '@pptx-editor.dev/react';
import { adapterHref } from '../../shared/demo-files.ts';

const DEV = import.meta.env.DEV;

/** Shown in development and in the preview build, which sets `VITE_ENABLE_FRAMEWORK_SWITCHER`. */
export const SHOW_ADAPTER_SWITCHER =
  DEV || import.meta.env.VITE_ENABLE_FRAMEWORK_SWITCHER === 'true';

const ADAPTERS = [
  { id: 'react', label: 'React', href: DEV ? 'http://localhost:5274/' : '/react/' },
  { id: 'vue', label: 'Vue', href: DEV ? 'http://localhost:5275/' : '/vue/' },
] as const;

export function AdapterSwitcher({ current }: { current: 'react' | 'vue' }) {
  const { t } = useTranslation();
  return (
    <nav className="demo-adapter-switcher" aria-label={t('app.adapter')}>
      {ADAPTERS.map((a) =>
        a.id === current ? (
          <span key={a.id} className="demo-adapter-switcher__option" aria-current="page">
            {a.label}
          </span>
        ) : (
          <a
            key={a.id}
            href={adapterHref(a.href, location.search)}
            className="demo-adapter-switcher__option"
            // The query at the click, so a theme or deck chosen since the render comes along.
            onClick={(e) =>
              e.currentTarget.setAttribute('href', adapterHref(a.href, location.search))
            }
          >
            {a.label}
          </a>
        )
      )}
    </nav>
  );
}
