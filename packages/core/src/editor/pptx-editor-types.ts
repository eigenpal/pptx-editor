// The public option and instance types of `createPptxEditor`, kept beside the facade so that
// `pptx-editor.ts` stays under the 1000-line cap.

import type {
  ClipboardPayload,
  CommandResult,
  Editor,
  EditorMode,
  EditorSelection,
  TableStyleChoice,
  ChartDataRead,
} from '../contracts/index.ts';
import type { ClipboardDrawingReader } from './clipboard-host.ts';
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
   * AutoCorrect while typing (default true): curly quotes, the ©, ®, and ™ signs, ellipses,
   * fractions, en dashes, sentence capitals, and lists started by typing `1.`, `1)`, `*`, or
   * `-` and a space.
   */
  readonly autoCorrect?: boolean;
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
  /**
   * Reads the DrawingML clipboard package from the system clipboard, for a host that can,
   * such as a desktop wrapper. Presentation applications write copied shapes in this package
   * (`DRAWING_CLIPBOARD_TYPES`) and as a picture, and browsers give a page only the picture.
   * When a paste holds a picture, the editor calls this reader and pastes the package's
   * shapes as editable shapes. Without a package, or without this option, the picture pastes.
   */
  readonly readClipboardDrawing?: ClipboardDrawingReader;
}

/**
 * Which measurer the open deck's layout runs on, and whether the host's fonts are still on
 * the way. Returned by {@link PptxEditorInstance.fontMeasurement}, as in DOCX Editor. @public
 */
export interface FontMeasurementState {
  /**
   * `fixed` estimates advance widths; `shaped` measures real font bytes with HarfBuzz. It is
   * `fixed` while no deck is open, and until the text engine and its faces arrive.
   */
  readonly measurer: 'fixed' | 'shaped';
  /**
   * True while font work for the deck is still running: the text engine, the packaged
   * substitutes, or faces the deck asked the `fonts` origins for. Late faces count for at most
   * 10 seconds. The deck shows with stand-in faces meanwhile, and `substitutedFonts` stays
   * empty, so a substitution notice does not show for fonts that are still loading.
   */
  readonly resolving: boolean;
  // No `producer`, unlike DOCX Editor: `DocumentSession.fontGeneration` identifies the
  // measurer here.
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
  /**
   * Which measurer the deck's layout runs on, and whether the host's fonts are still on the
   * way. A deck that opens waits at most 2 seconds for its fonts, then shows with stand-in
   * faces; the host's faces install when they arrive.
   */
  fontMeasurement(): FontMeasurementState;
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
  /**
   * The data of the selected chart, or of the chart frame `nodeId` on the current slide: its
   * categories, series names, and values, and where its workbook is. Refused with a reason
   * (a translation key) for a chart whose data shape the editor does not edit. The Edit Data
   * dialog shows it; `setChartData` writes it back.
   */
  chartData(nodeId?: string): ChartDataRead;
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
   * Open the rest of a deck that is still opening, now, for a reader of the whole deck (the
   * automation host's export). False when a slide did not parse. @internal
   */
  finishOpening(): boolean;
  /**
   * Publish a refusal as `lastRejection`, for chrome that refuses an action itself (Change
   * Picture after the selection changed while the file loaded). @internal
   */
  reject(reason: string, detail?: string): CommandResult;
}
