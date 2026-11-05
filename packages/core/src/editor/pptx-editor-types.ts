// The public option and instance types of `createPptxEditor`, kept beside the facade so that
// `pptx-editor.ts` stays under the 1000-line cap.

import type {
  ClipboardPayload,
  CommandResult,
  Editor,
  EditorMode,
  EditorSelection,
  TableStyleChoice,
} from '../contracts/index.ts';
import type { DocumentSession, OpenOptions } from './document-session.ts';
import type { EditorFontError } from './editor-font-error.ts';
import type { EditorFontsOption } from './host-fonts.ts';
import type { TextInputKind } from './text-input.ts';
import type { PptxFind } from './find-session.ts';
import type * as seam from './modules.ts';

/** Options for `createPptxEditor`. @public */
export interface PptxEditorOptions extends Omit<OpenOptions, 'fonts'> {
  readonly document?: Uint8Array;
  readonly mode?: EditorMode;
  readonly zoom?: number | 'fit';
  /** Whether drags snap to smart guides; on by default. `toggleSmartGuides` switches it. */
  readonly smartGuides?: boolean;
  /**
   * Translates engine-generated labels (placeholder frames, the canvas name). Adapters pass
   * their i18n translator; English keys show when omitted.
   */
  readonly translate?: (key: string, params?: Record<string, string | number>) => string;
  /**
   * Families the font picker offers besides the standard ones and the deck's own. A family
   * picked that no loaded face covers is requested from the `fonts` origins. `setFontFamilies`
   * replaces them.
   */
  readonly fontFamilies?: readonly string[];
  /**
   * Font bytes for the host's faces, as in DOCX Editor: a `FontConfiguration`, a
   * `FontConfigurationFragment` (`{ sources, substitutions }`), or a `FontResolver` such as
   * `packagedFonts()` or `googleFonts()` from `@docx-editor.dev/fonts`. A resolver is called
   * once per deck, after it opens, with the families the deck names (theme fonts first) and
   * its default family (the theme's minor Latin font), and again for a family picked in the
   * font picker. Host faces rank after the deck's embedded faces and before the packaged
   * substitutes. Omitted, the packaged substitutes load as before and nothing is fetched.
   * A resolver that fetches makes opening a deck perform network requests; the engine never
   * supplies one. Failures report through `onFontError`. `setFonts` replaces it.
   */
  readonly fonts?: EditorFontsOption;
  /** Called for each font failure: a failed origin, a refused face, or a failed resolution. */
  readonly onFontError?: (error: EditorFontError) => void;
  /** Capability modules, such as the review and collaboration modules of the pro package. */
  readonly modules?: readonly seam.EditorModule[];
}

/** The editor instance: the Editor contract plus engine extras. @public */
export interface PptxEditorInstance extends Editor {
  /**
   * Replace host font resources without replacing the document or its edit history, as DOCX
   * Editor's `setFonts`. The previous resolution is cancelled (its `AbortSignal` aborts).
   * `undefined` removes the host faces; a new value keeps the current faces until its faces
   * install. A value that supplies nothing usable, or fails, removes the replaced faces and
   * reports through `onFontError`. Embedded deck faces stay. The session, selection, undo
   * history, and collaboration session stay. Throws after `destroy()`.
   */
  setFonts(fonts: PptxEditorOptions['fonts']): Promise<void>;
  /**
   * The fonts in force: the value last passed as `fonts` or to `setFonts`. Compose it with
   * another origin (`withFontOrigin`) to add faces without dropping the host's.
   */
  readonly fonts: PptxEditorOptions['fonts'];
  /** The open session (store, measurer, layouts), for advanced hosts and tests. */
  readonly session: DocumentSession | null;
  /** Set the selection programmatically (shape ids or a text range). */
  setSelection(selection: EditorSelection): void;
  /** Families the font picker offers: standard, configured, embedded, and the deck's own. */
  fontFamilies(): readonly string[];
  /**
   * Replace the families the font picker offers besides the standard ones and the deck's own
   * (the `fontFamilies` option). A list with the same families, in the same order, changes
   * nothing; another list publishes a change. The document and the undo history stay.
   */
  setFontFamilies(families: readonly string[]): void;
  /** Table styles the Table Styles picker offers: the deck's own and the built-in default. */
  tableStyles(): readonly TableStyleChoice[];
  /** The shapes copied last in this editor (paste uses it when the clipboard is not readable). */
  readonly lastCopied: ClipboardPayload | null;
  /**
   * Apply text input at the current text selection, exactly as the keyboard does. A refusal
   * is returned and also published as `lastRejection`.
   */
  input(kind: TextInputKind, data?: string): CommandResult;
  /** Find and replace in slide text, tables, and speaker notes. */
  readonly find: PptxFind;
  /**
   * Paint how a new slide on a layout (`SlideLayoutSummary.id`) looks, with its placeholder
   * prompts, into a container: the Layout and New Slide galleries.
   */
  renderLayoutThumbnail(layoutId: string, container: HTMLElement): void;
  /**
   * Publish a refusal as `lastRejection`, for chrome that refuses an action itself (Change
   * Picture after the selection changed while the file loaded). @internal
   */
  reject(reason: string, detail?: string): CommandResult;
}
