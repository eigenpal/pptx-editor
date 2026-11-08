// After docx-editor/packages/core/src/__tests__/consumer.test-d.ts at 5993bd223.

/**
 * Consumer-side type test.
 *
 * The package typechecking on its own proves only that the declarations are internally
 * consistent. It does not prove a consumer can construct a command, read a result, or pass
 * the `fonts` option. Every check below imports a public entry point by its package name, the
 * way a host or an adapter writes it. Two commands compile this file:
 *
 * - `bun run typecheck`, through `tsconfig.consumer.json`, which maps the package names to the
 *   sources. Fast, and needs no build.
 * - `bun run check:consumer-types`, after `bun run build:packages`, which resolves the names
 *   through each `package.json` `exports` to the built `dist/*.d.ts`, as a host does. It sees
 *   breaks that exist only in the build (CI runs it in the build job).
 *
 * When this file fails to compile, the public types changed. `Unused '@ts-expect-error'
 * directive` means that a value the comment above it says is refused is now accepted. If the
 * change is intended, update the check here and add a changeset: `minor` for an addition,
 * `major` for a break.
 */

import {
  createPptxEditor,
  type PptxEditorInstance,
  type PptxEditorOptions,
} from '@pptx-editor.dev/core';
import {
  defineFontResolver,
  EditorFontError,
  withFontOrigin,
  type EditorFontsOption,
  type EditorModule,
} from '@pptx-editor.dev/core/editor';
import type {
  CanResult,
  CommandResult,
  Editor,
  EditorCommand,
  EditorSnapshot,
  FontConfiguration,
} from '@pptx-editor.dev/core/contracts';
import type { ResolvedSlide } from '@pptx-editor.dev/core/cascade';
import {
  layoutSlide,
  type FontConfigurationFragment,
  type FontResolutionRequest,
  type SlideLayout,
  type TextMeasurer,
} from '@pptx-editor.dev/core/layout';
import {
  PptxEditor as ReactPptxEditor,
  type PptxEditorProps as ReactPptxEditorProps,
  type PptxEditorRef as ReactPptxEditorRef,
  type PptxEditorRootProps as ReactPptxEditorRootProps,
} from '@pptx-editor.dev/react';
import {
  PptxEditor as VuePptxEditor,
  type PptxEditorProps as VuePptxEditorProps,
  type PptxEditorRef as VuePptxEditorRef,
  type PptxEditorRootProps as VuePptxEditorRootProps,
} from '@pptx-editor.dev/vue';

// A no-arg command must be constructible.
const undoCmd: EditorCommand = { type: 'undo' };
const boldCmd: EditorCommand = { type: 'toggleBold' };
// Commands with arguments.
const sizeCmd: EditorCommand = { type: 'setFontSize', points: 24 };
const shapeCmd: EditorCommand = { type: 'insertShape', preset: 'rect' };

// @ts-expect-error a command with arguments needs them
const sizeWithoutPoints: EditorCommand = { type: 'setFontSize' };

// @ts-expect-error an unknown command type is refused
const unknownCmd: EditorCommand = { type: 'notACommand' };

// The options take every documented form of `fonts`.
const fragment: FontConfigurationFragment = { sources: [], substitutions: [] };
const configuration: FontConfiguration = {
  epoch: 1,
  maxFontBytes: 1024,
  sources: [],
  defaultFont: { family: 'Calibri', sizeHalfPoints: 22 },
};
const resolver = defineFontResolver(async (request: FontResolutionRequest) => {
  void request;
  return fragment;
});
const fontsFromFragment: EditorFontsOption = fragment;
const fontsFromConfiguration: EditorFontsOption = configuration;
const fontsFromResolver: EditorFontsOption = resolver;

// @ts-expect-error `fonts` is not a list of family names
const fontsAsNames: EditorFontsOption = ['Calibri'];

const options: PptxEditorOptions = {
  document: new Uint8Array(),
  mode: 'view',
  zoom: 'fit',
  fonts: resolver,
  onFontError: (error) => {
    const typed: EditorFontError = error;
    const code: string = typed.code;
    void code;
  },
};

// @ts-expect-error `zoom` is a number or 'fit'
const badZoom: PptxEditorOptions = { zoom: 'page' };

declare const resolvedSlide: ResolvedSlide;

export function exercise(measurer: TextMeasurer): void {
  const editor: PptxEditorInstance = createPptxEditor(options);
  // The instance IS the `Editor` contract.
  const contract: Editor = editor;
  void contract;

  // Writes return a result that narrows.
  const result: CommandResult = editor.exec(boldCmd);
  if (result.ok) {
    // @ts-expect-error a success carries no reason
    void result.reason;
  } else {
    const reason: string = result.reason;
    const noop: boolean | undefined = result.noop;
    void reason;
    void noop;
  }

  // `can` is a dry run that names its reason when it refuses.
  const dry: CanResult = editor.can(undoCmd);
  if (!dry.ok) {
    const reason: string = dry.reason;
    void reason;
  }

  // Snapshot payloads carry real shapes.
  const snap: EditorSnapshot = editor.snapshot();
  const status: 'empty' | 'loading' | 'ready' | 'error' = snap.status;
  const slide: number = snap.currentSlide;
  void status;
  void slide;

  // Fonts replace in place, and compose with another origin.
  void editor.setFonts(withFontOrigin(editor.fonts, fragment));
  void editor.setFonts(undefined);
  const saved: Uint8Array | null = editor.save();
  void saved;
  // The saved bytes are a `BlobPart` as they are: TypeScript 5.7+ refuses a
  // `Uint8Array<ArrayBufferLike>` there, so `save()` returns `Uint8Array<ArrayBuffer>`.
  const bytes = editor.save();
  if (bytes) void new Blob([bytes]);
  // The font picker's host families replace in place.
  editor.setFontFamilies(['Inter']);

  // The layout lane is usable on its own, over a resolved slide and an injected measurer.
  const layout: SlideLayout = layoutSlide(resolvedSlide, measurer);
  void layout;
  editor.destroy();
}

// The React sugar component takes the same `fonts` option and hands out the instance.
const reactProps: ReactPptxEditorProps = {
  document: new Uint8Array(),
  mode: 'edit',
  fonts: resolver,
  colorMode: 'system',
  onReady: (editor) => {
    const instance: PptxEditorInstance = editor;
    void instance;
  },
};
const reactRootProps: ReactPptxEditorRootProps = { fonts: fragment, fontFamilies: ['Inter'] };
// Both adapters take capability modules, as `createPptxEditor({ modules })` does.
declare const reviewModule: EditorModule;
const reactModules: ReactPptxEditorProps = { modules: [reviewModule] };
const reactRootModules: ReactPptxEditorRootProps = { modules: [reviewModule] };
const vueModules: VuePptxEditorProps = { modules: [reviewModule] };
const vueRootModules: VuePptxEditorRootProps = { modules: [reviewModule] };
// The adapter refs hand out the same `BlobPart`-ready bytes.
declare const reactRef: ReactPptxEditorRef;
declare const vueRef: VuePptxEditorRef;
const reactSaved = reactRef.save();
const vueSaved = vueRef.save();
if (reactSaved && vueSaved) void new Blob([reactSaved, vueSaved]);
// @ts-expect-error `mode` is 'edit' or 'view'
const reactBadMode: ReactPptxEditorProps = { mode: 'present' };
const reactRoot: unknown = ReactPptxEditor.Root;

// The Vue sugar component's props derive from its runtime props, and take the same `fonts`.
const vueProps: VuePptxEditorProps = { document: new Uint8Array(), mode: 'view', fonts: resolver };
// @ts-expect-error `mode` is 'edit' or 'view'
const vueBadMode: VuePptxEditorProps = { mode: 'present' };
const vueName: string | undefined = VuePptxEditor.name;

void shapeCmd;
void sizeCmd;
void sizeWithoutPoints;
void unknownCmd;
void fontsFromFragment;
void fontsFromConfiguration;
void fontsFromResolver;
void fontsAsNames;
void badZoom;
void reactProps;
void reactRootProps;
void reactModules;
void reactRootModules;
void vueModules;
void vueRootModules;
void reactBadMode;
void reactRoot;
void vueProps;
void vueBadMode;
void vueName;
