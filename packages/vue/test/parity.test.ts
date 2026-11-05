// The parity contract against the real adapters, and each kind of drift it must catch. Runs in
// the DOM test run, because the Vue side is read by mounting the components.

import { beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { parityErrors } from '../../../scripts/check-parity-contract.mjs';
import { vueAdapterShape, type VueAdapterShape } from './parity-shape.ts';

const root = join(import.meta.dir, '../../..');
const read = (path: string): string => readFileSync(join(root, path), 'utf8');

let shape: VueAdapterShape;
beforeAll(async () => {
  shape = await vueAdapterShape(JSON.parse(read('scripts/parity/parity.contract.json')));
});

const inputs = () => ({
  reactApi: read('docs/api/react/index.api.md'),
  vueApi: read('docs/api/vue/index.api.md'),
  vue: shape,
  contract: JSON.parse(read('scripts/parity/parity.contract.json')),
});

describe('check:parity-contract', () => {
  test('the committed adapters match the contract', () => {
    expect(parityErrors(inputs())).toEqual([]);
  });

  test('the probes see the class and the slots of the paired components', () => {
    expect(shape.PptxEditor.classLands).toBe(true);
    expect(shape.PptxEditor.slots).toContain('default');
    expect(shape.PptxEditorToolbar.classLands).toBe(true);
    expect(shape.PptxEditorToolbar.classNameLands).toBe(true);
    expect(shape.PptxEditorRoot.slots).toContain('default');
    // A named slot: the loading screen's indicator renders when there is no default content.
    expect(shape.PptxEditorLoading.slots).toEqual(['default', 'indicator']);
    // The link popup teleports to the page body; the probe reads the document.
    expect(shape.HyperlinkPopup.classLands).toBe(true);
  });

  test('a prop inherited through extends must be classified', () => {
    const input = inputs();
    input.reactApi = input.reactApi.replace(
      /export interface PaneProps \{/,
      'export interface PaneProps {\n    readonly zoom?: number;'
    );
    const errors = parityErrors(input);
    expect(errors).toContain(
      'PptxEditorSlideRail (SlideRailProps): React prop "zoom" is not classified in the contract.'
    );
    expect(errors).toContain(
      'PptxEditorNotesPane (NotesPaneProps): React prop "zoom" is not classified in the contract.'
    );
  });

  test('a prop removed with Omit<> is not a React prop', () => {
    const input = inputs();
    input.reactApi = input.reactApi.replace(
      "export interface NotesPaneProps extends Omit<PaneProps, 'paneWidth'> {",
      'export interface NotesPaneProps extends PaneProps {'
    );
    expect(parityErrors(input)).toContain(
      'PptxEditorNotesPane (NotesPaneProps): React prop "paneWidth" is not classified in the contract.'
    );
  });

  test('a base that is not in the snapshot fails', () => {
    const input = inputs();
    input.reactApi = input.reactApi.replace(
      'export interface FontDialogProps extends DialogCustomizationProps {',
      'export interface FontDialogProps extends HiddenBase {'
    );
    expect(parityErrors(input)).toContain(
      'FontDialog (FontDialogProps): FontDialogProps extends HiddenBase, which is not an interface of the snapshot; its members cannot be checked.'
    );
  });

  test('a paired prop whose type differs between the adapters fails', () => {
    const input = inputs();
    input.reactApi = input.reactApi.replace(
      /(export interface HyperlinkPopupProps \{[^}]*?readonly url: )string;/,
      '$1URL;'
    );
    expect(parityErrors(input)).toContain(
      'HyperlinkPopup (HyperlinkPopupProps): paired "url" type drift: React `URL`, Vue `string`.'
    );
  });

  test('a Vue emit that the contract does not map fails', () => {
    const input = inputs();
    const rail = shape.PptxEditorSlideRail;
    input.vue = { ...shape, PptxEditorSlideRail: { ...rail, emits: [...rail.emits, 'resize'] } };
    expect(parityErrors(input)).toContain(
      'PptxEditorSlideRail (SlideRailProps): Vue emit "resize" is not declared in reactCallbacksAsVueEmits.'
    );
  });

  test('a named slot that does not render fails', () => {
    const input = inputs();
    input.vue = { ...shape, PptxEditorLoading: { ...shape.PptxEditorLoading, slots: ['default'] } };
    expect(parityErrors(input)).toContain(
      'PptxEditorLoading (LoadingProps): Vue PptxEditorLoading does not render its "indicator" slot.'
    );
  });

  test('a className that does not reach an element fails', () => {
    const input = inputs();
    const pill = shape.PptxEditorZoomPill;
    input.vue = { ...shape, PptxEditorZoomPill: { ...pill, classNameLands: false } };
    expect(parityErrors(input)).toContain(
      'PptxEditorZoomPill (ZoomPillProps): the className prop of Vue PptxEditorZoomPill does not reach a rendered element.'
    );
  });

  test('a Vue part that ignores a paired prop value fails', () => {
    expect(shape.PptxEditorSlideRail.values).toEqual({ orientation: true, paneWidth: true });
    const input = inputs();
    const rail = shape.PptxEditorSlideRail;
    input.vue = {
      ...shape,
      PptxEditorSlideRail: { ...rail, values: { orientation: true, paneWidth: false } },
    };
    expect(parityErrors(input)).toContain(
      'PptxEditorSlideRail (SlideRailProps): Vue PptxEditorSlideRail ignores a value of its paired "paneWidth" prop.'
    );
  });

  test('stale entries fail', () => {
    const input = inputs();
    input.contract.components.PptxEditorSidePanel.vueExclusive = { gone: 'Removed.' };
    input.contract.components.PptxEditorSidePanel.deferredInVue = { gone2: 'Removed.' };
    input.contract.exports.vueExclusive.NoSuchExport = 'Removed.';
    const errors = parityErrors(input);
    expect(errors).toContain(
      'PptxEditorSidePanel (SidePanelProps): vueExclusive "gone" is missing in Vue (contract stale).'
    );
    expect(errors).toContain(
      'PptxEditorSidePanel (SidePanelProps): deferredInVue "gone2" is missing in React (contract stale).'
    );
    expect(errors).toContain(
      'exports: vueExclusive "NoSuchExport" is missing in Vue (contract stale).'
    );
  });

  test('a React component with a props type needs an entry or a propsUnchecked reason', () => {
    const input = inputs();
    input.reactApi +=
      '\nexport function NewPart(input: NewPartProps): react.JSX.Element;\n' +
      'export interface NewPartProps {\n    readonly className?: string;\n}\n';
    expect(parityErrors(input)).toContain(
      'components: React component "NewPart" has no components entry and no propsUnchecked reason.'
    );
    input.contract.propsUnchecked.NewPart = 'Probe-only part.';
    expect(parityErrors(input).some((e) => e.includes('component "NewPart"'))).toBe(false);
  });

  test('the top-level PptxEditor takes callbacks as emits, never as props', () => {
    const input = inputs();
    const editor = shape.PptxEditor;
    input.vue = { ...shape, PptxEditor: { ...editor, props: [...editor.props, 'onReady'] } };
    expect(parityErrors(input)).toContain(
      'PptxEditor (PptxEditorProps): "onReady" must be a Vue emit, not a prop.'
    );
  });

  test('a new React prop with no contract entry fails', () => {
    const input = inputs();
    input.reactApi = input.reactApi.replace(
      /export interface PptxEditorProps \{/,
      'export interface PptxEditorProps {\n    readonly zoom?: number;'
    );
    expect(parityErrors(input)).toContain(
      'PptxEditor (PptxEditorProps): React prop "zoom" is not classified in the contract.'
    );
  });

  test('a renamed export missing in Vue fails', () => {
    const input = inputs();
    input.vueApi = input.vueApi.replace(/^export const PptxEditorToolbar\b.*$/m, '');
    expect(parityErrors(input)).toContain(
      'exports: Vue has no "PptxEditorToolbar" (pair of "Toolbar").'
    );
  });

  test('a class that does not reach an element fails', () => {
    const input = inputs();
    input.vue = { ...shape, PptxEditor: { ...shape.PptxEditor, classLands: false } };
    expect(parityErrors(input)).toContain(
      'PptxEditor (PptxEditorProps): a class on Vue PptxEditor does not reach a rendered element.'
    );
  });

  test('a deferred prop that ships in Vue must move to paired', () => {
    const input = inputs();
    input.reactApi = input.reactApi.replace(
      /export interface PptxEditorProps \{/,
      'export interface PptxEditorProps {\n    readonly zoom?: number;'
    );
    input.contract.components.PptxEditor.deferredInVue = { zoom: 'Not built in Vue yet.' };
    input.vue = {
      ...shape,
      PptxEditor: { ...shape.PptxEditor, props: [...shape.PptxEditor.props, 'zoom'] },
    };
    expect(parityErrors(input)).toContain(
      'PptxEditor (PptxEditorProps): deferredInVue "zoom" has shipped in Vue; move it to paired.'
    );
  });

  test('a React callback without a Vue emit fails', () => {
    const input = inputs();
    input.vue = { ...shape, PptxEditorRoot: { ...shape.PptxEditorRoot, emits: ['ready'] } };
    expect(parityErrors(input)).toContain(
      'PptxEditorRoot (PptxEditorRootProps): Vue PptxEditorRoot does not emit "change".'
    );
  });

  test('a React-only export that appears in Vue must be reclassified', () => {
    const input = inputs();
    input.vueApi += '\nexport function translatorFor(): void;\n';
    expect(parityErrors(input)).toContain(
      'exports: reactExclusive "translatorFor" appeared in Vue; reclassify it.'
    );
  });

  test('a pair with a different release tag in the two adapters fails', () => {
    const input = inputs();
    input.vueApi = input.vueApi.replace(
      /\/\/ @internal\r?\nexport const ToolButton\b/,
      '// @public\nexport const ToolButton'
    );
    expect(parityErrors(input)).toContain(
      'exports: "ToolButton" is @internal in React but "ToolButton" is @public in Vue.'
    );
  });
});

describe('check:parity-contract ref gate', () => {
  const REF = [
    '    exec(command: EditorCommand, options?: EditorExecOptions): CommandResult;',
    '    focus(): void;',
    '    getEditor(): PptxEditorInstance | null;',
    '    load(bytes: Uint8Array): Promise<CommandResult>;',
    '    save(): Uint8Array | null;',
    '    snapshot(): EditorSnapshot;',
  ];
  // API Extractor writes CRLF line endings.
  const HEADER = /\r?\nexport interface PptxEditorRef \{[\s\S]*?\r?\n\}\r?\n/;
  /** The snapshot with its `PptxEditorRef` block replaced by one with `members`. */
  const withRef = (api: string, members: readonly string[]): string =>
    `${api.replace(HEADER, '\n')}\n// @public\nexport interface PptxEditorRef {\n${members.join('\n')}\n}\n`;
  const refErrors = (react: readonly string[], vue: readonly string[]): string[] => {
    const input = inputs();
    input.reactApi = withRef(input.reactApi, react);
    input.vueApi = withRef(input.vueApi, vue);
    return parityErrors(input).filter((e) => e.startsWith('ref:'));
  };

  test('the same members on both refs pass', () => {
    expect(refErrors(REF, REF)).toEqual([]);
  });

  test('a member on one ref only fails', () => {
    expect(refErrors([...REF, '    print(): void;'], REF)).toContain(
      'ref: React PptxEditorRef member "print" is not classified in the contract.'
    );
    expect(refErrors(REF, REF.slice(1))).toContain('ref: paired "exec" is missing in Vue.');
  });

  test('a paired member with another type fails', () => {
    const vue = REF.map((m) => (m.includes('save()') ? '    save(): Uint8Array;' : m));
    expect(refErrors(REF, vue)).toContain(
      'ref: paired "save" type drift: React `() => Uint8Array | null`, Vue `() => Uint8Array`.'
    );
  });

  test('a snapshot without the interface fails', () => {
    const input = inputs();
    input.vueApi = input.vueApi.replace(HEADER, '\n');
    expect(parityErrors(input)).toContain('ref: Vue snapshot has no PptxEditorRef interface.');
  });
});
