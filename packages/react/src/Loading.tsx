/** @jsxImportSource react */
// The loading screen, after DOCX Editor's `DocxEditor.Loading`: what the editor shows while it
// has no deck to paint yet, and while a deck opens. It reads the same snapshot as every other
// part (`showsLoading` in core), so the screen and the chrome never disagree, and it renders
// nothing once the deck is ready. The Vue twin is `PptxEditorLoading`.
//
// The engine publishes `loading` and then yields one frame before the parse, which blocks the
// main thread for a large deck, so this screen paints first.
//
// It carries `pptx-editor` itself, as `PptxEditor.Viewport` does, so the `--doc-*` tokens
// resolve wherever it is composed.

import type { CSSProperties, ReactNode } from 'react';
import { loadingSheetSize, sameLoadingSheetSize, showsLoading } from '@pptx-editor.dev/core/editor';
import { useChromeDark } from './color-mode.ts';
import { useEditorState, useTranslation } from './context.ts';

/** Props for `PptxEditor.Loading`. @public */
export interface LoadingProps {
  /**
   * An extra host condition, OR-ed with the editor's own. The default already shows the screen
   * while the editor has no deck, so pass this only for work the editor cannot see, such as a
   * host that mounts the root after its own fetch.
   */
  readonly when?: boolean;
  /**
   * Pin the screen over the nearest positioned ancestor, so it covers the previous deck while
   * the next one opens. A short appearance delay keeps a fast open from flashing it. Without
   * `overlay`, the screen is an in-flow box that fills its parent.
   */
  readonly overlay?: boolean;
  /** Appended after the `pptx-editor pptx-loading` classes. */
  readonly className?: string;
  /** Inline styles for the loading container. */
  readonly style?: CSSProperties;
  /**
   * Your own indicator in place of the packaged spinner, centered on the slide-shaped sheet.
   * Include text for screen readers, as the packaged indicator does.
   */
  readonly indicator?: ReactNode;
  /** The whole loading screen in place of the packaged sheet. */
  readonly children?: ReactNode;
}

/** Props for `PptxEditor.Loading.Spinner`. @public */
export interface LoadingSpinnerProps {
  /** Appended after the `pptx-loading__spinner` class. */
  readonly className?: string;
}

/**
 * The packaged spinner on its own, to compose into custom `children`. It is decorative
 * (`aria-hidden`), so give the surrounding screen its own text. @public
 */
export function LoadingSpinner({ className }: LoadingSpinnerProps) {
  return (
    <span
      className={`pptx-loading__spinner${className ? ` ${className}` : ''}`}
      aria-hidden="true"
    />
  );
}

function LoadingImpl({
  when = false,
  overlay = false,
  className,
  style,
  indicator,
  children,
}: LoadingProps) {
  const loading = useEditorState(showsLoading);
  const sheet = useEditorState(loadingSheetSize, sameLoadingSheetSize);
  const dark = useChromeDark();
  const { t } = useTranslation();
  if (!when && !loading) return null;
  const classes = ['pptx-editor', dark ? 'dark' : '', 'pptx-loading'];
  if (overlay) classes.push('pptx-loading--overlay');
  if (className) classes.push(className);
  return (
    <div
      className={classes.filter(Boolean).join(' ')}
      style={style}
      role="status"
      aria-live="polite"
    >
      {children ?? (
        <div
          className="pptx-loading__sheet"
          style={{ width: sheet.width, aspectRatio: `${sheet.width} / ${sheet.height}` }}
        >
          {indicator ?? (
            <>
              <LoadingSpinner />
              <span className="pptx-sr-only">{t('app.loading')}</span>
            </>
          )}
        </div>
      )}
    </div>
  );
}

/** The loading part, with the packaged spinner as a static. @public */
export interface LoadingComponent {
  /** Renders the loading screen, or nothing once the deck is ready. */
  (props: LoadingProps): ReactNode;
  /** The packaged indicator, to compose into custom children. */
  readonly Spinner: typeof LoadingSpinner;
}

/**
 * Shows a slide-shaped sheet with a spinner while the editor has no deck, and while a deck
 * opens; renders nothing once the deck is ready. A parse failure clears it, so a broken file
 * never spins forever: report it from `snapshot().error`. Compose it inside a positioned box
 * with `overlay` to cover the previous deck while the next one opens:
 *
 * ```tsx
 * <PptxEditor.Root document={bytes}>
 *   <PptxEditor.Viewport>
 *     <div style={{ position: 'relative', flex: 1 }}>
 *       <PptxEditor.Content />
 *       <PptxEditor.Loading overlay />
 *     </div>
 *   </PptxEditor.Viewport>
 * </PptxEditor.Root>
 * ```
 *
 * Outside a `PptxEditor.Root` it always shows, because no editor reports otherwise.
 *
 * @public
 */
export const Loading: LoadingComponent = Object.assign(LoadingImpl, { Spinner: LoadingSpinner });
