import type { PptxEditorInstance } from './pptx-editor-types.ts';
import { geometryOf } from './selection-state.ts';
import { itemById } from './surface.ts';
import { currentSlideSummary } from './master-rail.ts';
import type { Editor, EditorSnapshot, LineEndType, ShapeSummary } from '../contracts/index.ts';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import type { ChromeSlotId } from './chrome-controls.ts';

/** Stable identifiers for the selection's formatting sections. @public */
export type FormatSectionId =
  | 'fill'
  | 'line'
  | 'effects'
  | 'text'
  | 'paragraph'
  | 'picture'
  | 'table'
  | 'chart'
  | 'diagram'
  | 'equation'
  | 'arrange'
  | 'geometry'
  | 'slide';

/** A built-in formatting section. Commands retain their own capability checks. @public */
export interface FormatSectionDefinition {
  readonly id: FormatSectionId;
  readonly labelKey: TranslationKey;
  readonly slots: readonly ChromeSlotId[];
}

/** The default section order, shared by both adapters. @public */
export const FORMAT_SECTIONS: readonly FormatSectionDefinition[] = [
  { id: 'fill', labelKey: 'toolbar.fill', slots: ['shape.fill', 'shape.styles'] },
  {
    id: 'line',
    labelKey: 'toolbar.lineColor',
    slots: [
      'shape.lineColor',
      'shape.lineWidth',
      'shape.lineDash',
      'shape.lineStart',
      'shape.lineEnd',
    ],
  },
  {
    id: 'effects',
    labelKey: 'toolbar.shapeEffects',
    slots: ['shape.effects', 'shape.textEffects'],
  },
  {
    id: 'picture',
    labelKey: 'toolbar.groups.image',
    slots: [
      'image.change',
      'image.crop',
      'image.cropToShape',
      'image.cropFit',
      'image.corrections',
      'image.color',
      'image.transparency',
      'image.reset',
      'image.altText',
    ],
  },
  {
    id: 'text',
    labelKey: 'toolbar.groups.font',
    slots: [
      'font.family',
      'font.size',
      'text.bold',
      'text.italic',
      'text.underline',
      'text.color',
      'text.highlight',
      'text.dialog',
      'text.link',
    ],
  },
  {
    id: 'paragraph',
    labelKey: 'toolbar.groups.paragraph',
    slots: [
      'paragraph.align',
      'paragraph.valign',
      'list.bulletStyle',
      'list.numberStyle',
      'list.lineSpacing',
      'paragraph.textDirection',
      'paragraph.autofit',
      'paragraph.columns',
      'paragraph.dialog',
    ],
  },
  {
    id: 'table',
    labelKey: 'panel.format.table',
    slots: [
      'table.style',
      'table.cellFill',
      'table.borderTarget',
      'table.borderColor',
      'table.borderStyle',
      'table.borderWidth',
      'table.cellMargins',
      'table.insertRowAbove',
      'table.insertRowBelow',
      'table.insertColumnLeft',
      'table.insertColumnRight',
      'table.mergeCells',
      'table.splitCell',
      'table.distributeRows',
      'table.distributeColumns',
      'table.delete',
    ],
  },
  { id: 'chart', labelKey: 'panel.tabs.chart', slots: [] },
  { id: 'diagram', labelKey: 'panel.tabs.diagram', slots: [] },
  { id: 'equation', labelKey: 'toolbar.equation', slots: ['insert.equation'] },
  {
    id: 'arrange',
    labelKey: 'toolbar.groups.arrange',
    slots: ['arrange.order', 'arrange.align', 'arrange.rotate', 'arrange.group', 'arrange.ungroup'],
  },
  { id: 'geometry', labelKey: 'inspector.geometry', slots: [] },
  { id: 'slide', labelKey: 'panel.format.slide', slots: ['slide.layout', 'slide.background'] },
];

/** The selected object's editing category. @public */
export type FormatSelectionCategory =
  | 'line'
  | 'shape'
  | 'text'
  | 'picture'
  | 'table'
  | 'chart'
  | 'diagram'
  | 'equation'
  | 'group'
  | 'slide'
  | 'mixed';

function selectedSummaries(snapshot: EditorSnapshot): readonly (ShapeSummary | undefined)[] {
  const flatten = (items: readonly ShapeSummary[]): readonly ShapeSummary[] =>
    items.flatMap((item) => [item, ...flatten(item.children ?? [])]);
  const summaries = new Map(flatten(snapshot.shapes ?? []).map((s) => [s.nodeId, s]));
  return snapshot.selection.kind === 'shapes'
    ? snapshot.selection.nodeIds.map((id) => summaries.get(id))
    : [];
}

/** True only when every selected object supports endpoint types. @internal */
export function formatSelectionHasLineEnds(snapshot: EditorSnapshot): boolean {
  if (snapshot.geometry?.lineHead !== undefined) return true;
  const chosen = selectedSummaries(snapshot);
  return chosen.length > 0 && chosen.every((s) => s?.lineEnds === true || s?.kind === 'connector');
}

/** Determine primary inspector tasks without changing document selection. @public */
export function formatSelectionCategory(snapshot: EditorSnapshot): FormatSelectionCategory {
  if (snapshot.selection.kind === 'none') return 'slide';
  if (snapshot.table) return 'table';
  if (snapshot.chart) return 'chart';
  if (snapshot.diagram) return 'diagram';
  if (snapshot.equation) return 'equation';
  if (snapshot.selection.kind === 'text') return 'text';
  if (snapshot.image) return 'picture';
  if (formatSelectionHasLineEnds(snapshot)) return 'line';
  const chosen = selectedSummaries(snapshot);
  const kind =
    snapshot.geometry?.kind ??
    (chosen.length && chosen.every((s) => s?.kind === chosen[0]?.kind)
      ? chosen[0]?.kind
      : undefined);
  return kind === 'shape' || kind === 'picture' || kind === 'group' ? kind : 'mixed';
}

const PRIMARY: Readonly<Record<FormatSelectionCategory, readonly FormatSectionId[]>> = {
  line: ['line'],
  shape: ['fill', 'line'],
  text: ['text', 'paragraph'],
  picture: ['picture'],
  table: ['table'],
  chart: ['chart'],
  diagram: ['diagram'],
  equation: ['equation'],
  group: ['arrange', 'geometry'],
  slide: ['slide'],
  mixed: ['line', 'fill'],
};
interface ExpansionState {
  values: Map<string, boolean>;
  listeners: Set<() => void>;
}
const expansion = new WeakMap<Editor, ExpansionState>();
function expansionOf(editor: Editor): ExpansionState {
  let state = expansion.get(editor);
  if (!state) {
    state = { values: new Map(), listeners: new Set() };
    expansion.set(editor, state);
  }
  return state;
}

/** Read per-editor disclosure preferences, or the category's defaults. No document state is written. @public */
export function formatSectionOpen(
  editor: Editor | null,
  category: FormatSelectionCategory,
  id: FormatSectionId
): boolean {
  const selected = editor ? expansionOf(editor).values.get(`${category}.${id}`) : undefined;
  if (selected !== undefined) return selected;
  if (id === 'effects') {
    const effects = editor?.snapshot().geometry?.effects;
    if (
      effects &&
      (effects.shadow !== 'none' ||
        effects.reflection !== 'none' ||
        effects.glow > 0 ||
        effects.softEdge > 0 ||
        effects.blur > 0 ||
        effects.bevel !== 'none' ||
        effects.rotation3d !== 'none')
    )
      return true;
  }
  return PRIMARY[category].includes(id);
}

/** Store a disclosure preference for one selection category. @public */
export function setFormatSectionOpen(
  editor: Editor | null,
  category: FormatSelectionCategory,
  id: FormatSectionId,
  open: boolean
): void {
  if (!editor) return;
  const state = expansionOf(editor),
    key = `${category}.${id}`;
  if (state.values.get(key) === open) return;
  state.values.set(key, open);
  for (const listener of state.listeners) listener();
}

/** Observe disclosure preferences. Call the returned function when the inspector unmounts. @public */
export function listenFormatSectionExpansion(
  editor: Editor | null,
  listener: () => void
): () => void {
  if (!editor) return () => undefined;
  const state = expansionOf(editor);
  state.listeners.add(listener);
  return () => {
    state.listeners.delete(listener);
  };
}

/** Common outline values. Null values are mixed or unavailable, never a numeric default. @public */
export interface FormatLineState {
  readonly width: number | null;
  readonly widthMixed: boolean;
  readonly head: LineEndType | null;
  readonly tail: LineEndType | null;
  readonly headMixed: boolean;
  readonly tailMixed: boolean;
  readonly endsSupported: boolean;
}

/** Read exact points and endpoint types across the selection. @public */
export function formatLineState(editor: Editor | null): FormatLineState | null {
  if (!editor) return null;
  const snapshot = editor.snapshot();
  const session = (editor as Partial<PptxEditorInstance>).session;
  const layout = session?.layout(snapshot.currentSlide);
  const geometries =
    snapshot.selection.kind === 'shapes' && snapshot.selection.nodeIds.length > 1
      ? snapshot.selection.nodeIds.map((id) =>
          geometryOf(layout, { kind: 'shapes', nodeIds: [id] })
        )
      : [snapshot.geometry];
  if (
    !geometries.length ||
    geometries.some((g) => !g || !['shape', 'connector', 'picture'].includes(g.kind))
  )
    return null;
  // Public geometry keeps its visible-only, rounded toolbar value. Read resolved lines
  // independently so imported EMU precision and mixed widths survive the inspector.
  const widths = geometries.map((g) => {
    const element = layout && itemById(layout, g!.nodeId)?.element;
    return element && 'line' in element
      ? element.line
        ? element.line.width / 12_700
        : undefined
      : g!.lineWidth;
  });
  const heads = geometries.map((g) => g!.lineHead);
  const tails = geometries.map((g) => g!.lineTail);
  const mixed = (values: readonly unknown[]) => values.some((v) => v !== values[0]);
  const endsSupported = geometries.every((g) => g!.lineHead !== undefined);
  return {
    width: mixed(widths) ? null : (widths[0] ?? null),
    widthMixed: mixed(widths),
    head: endsSupported && !mixed(heads) ? heads[0]! : null,
    tail: endsSupported && !mixed(tails) ? tails[0]! : null,
    headMixed: endsSupported && mixed(heads),
    tailMixed: endsSupported && mixed(tails),
    endsSupported,
  };
}

/** Applicable sections, optionally ordered and filtered by the host. Omit an id to hide it. @public */
export function formatSectionsOf(
  snapshot: EditorSnapshot,
  order?: readonly FormatSectionId[]
): readonly FormatSectionDefinition[] {
  const { geometry: g, selection: s } = snapshot;
  const selected = s.kind !== 'none';
  const shape = g?.kind === 'shape';
  const chosen = selectedSummaries(snapshot);
  const multiple = chosen.length > 1;
  const commonFill = multiple && chosen.every((shape) => shape?.kind === 'shape');
  const commonOutline =
    multiple &&
    chosen.every((shape) => shape && ['shape', 'picture', 'connector'].includes(shape.kind));
  const text = !!snapshot.format && (!multiple || commonFill);
  const visible: Record<FormatSectionId, boolean> = {
    fill: shape || commonFill,
    line: (!!g && ['shape', 'picture', 'connector'].includes(g.kind)) || commonOutline,
    effects: (!!g && ['shape', 'picture', 'connector'].includes(g.kind)) || commonOutline,
    text,
    paragraph: text,
    picture: !!snapshot.image,
    table: !!snapshot.table,
    chart: !!snapshot.chart,
    diagram: !!snapshot.diagram,
    equation: !!snapshot.equation,
    arrange: selected,
    geometry: !!g,
    slide: !selected,
  };
  const ids = order ?? [
    ...PRIMARY[formatSelectionCategory(snapshot)],
    ...FORMAT_SECTIONS.map((s) => s.id),
  ];
  return [...new Set(ids)].flatMap((id) => {
    const section = FORMAT_SECTIONS.find((s) => s.id === id);
    if (!section || !visible[id]) return [];
    const slots = section.slots.filter((slot) => {
      if (slot === 'shape.lineStart' || slot === 'shape.lineEnd')
        return formatSelectionHasLineEnds(snapshot);
      if (slot === 'shape.textEffects') return text;
      return true;
    });
    const labelKey =
      id === 'line'
        ? formatSelectionHasLineEnds(snapshot)
          ? 'inspector.line'
          : 'inspector.outline'
        : section.labelKey;
    return slots.length === section.slots.length && labelKey === section.labelKey
      ? [section]
      : [{ ...section, labelKey, slots }];
  });
}

/** Identifies the selection scope for panel drafts. @internal */
export function formatTargetKey(s: EditorSnapshot): string {
  return JSON.stringify([s.status, s.currentSlide, currentSlideSummary(s)?.id, s.selection]);
}

/** Whether a supported formatting control has different selected object values. @internal */
export function formatControlMixed(editor: PptxEditorInstance | null, slot: ChromeSlotId): boolean {
  if (!editor) return false;
  const snapshot = editor.snapshot();
  const font = snapshot.format?.font?.disagrees;
  if (slot === 'shape.lineWidth') return formatLineState(editor)?.widthMixed === true;
  if (slot === 'font.family') return font?.fontFamily === true;
  if (slot === 'font.size') return font?.fontSizePt === true;
  if (slot === 'text.bold') return font?.bold === true;
  if (slot === 'text.italic') return font?.italic === true;
  if (slot === 'text.underline') return font?.underline === true;
  if (snapshot.selection.kind !== 'shapes' || snapshot.selection.nodeIds.length < 2) return false;
  const layout = editor.session?.layout(snapshot.currentSlide);
  const values = snapshot.selection.nodeIds
    .map((id) => {
      const g = geometryOf(layout, { kind: 'shapes', nodeIds: [id] });
      if (!g) return undefined;
      switch (slot) {
        case 'shape.fill':
          return [g.fill, g.gradient, g.pattern];
        case 'shape.lineColor':
          return g.lineColor;
        case 'shape.lineStart':
          return g.lineHead;
        case 'shape.lineEnd':
          return g.lineTail;
        case 'shape.lineDash':
          return g.lineDash;
        case 'shape.effects':
          return g.effects;
        default:
          return undefined;
      }
    })
    .map((value) => JSON.stringify(value));
  return values.some((value) => value !== values[0]);
}

/** The common precise value for a line control; mixed values remain undefined. @internal */
export function formatLineSlotValue(editor: Editor | null, slot: ChromeSlotId): string | undefined {
  if (!['shape.lineWidth', 'shape.lineStart', 'shape.lineEnd'].includes(slot)) return undefined;
  const state = formatLineState(editor);
  if (!state) return undefined;
  const value =
    slot === 'shape.lineWidth' ? state.width : slot === 'shape.lineStart' ? state.head : state.tail;
  return value === null ? undefined : String(value);
}
