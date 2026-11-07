// The editing chrome of empty placeholders on the slide: a dotted outline shows where each one
// is, and an empty content, table, or picture placeholder shows insert buttons in its middle.
// A press on a button selects the placeholder and inserts a table or asks for a picture; the
// new object takes the placeholder's place (`automation/commands/placeholder-fill.ts`). It is
// editing chrome: the slide paint, the slide show, thumbnails, and saved files never show it.

import type { LayoutItem, SlideLayout } from '../layout/index.ts';
import type { EditorCommand, EditorSelection } from '../contracts/index.ts';
import { SVG_NS } from '../output/index.ts';
import { ICON_PATHS } from './icon-paths.ts';
import {
  PICTURE_PLACEHOLDERS,
  TABLE_PLACEHOLDERS,
} from '../automation/commands/placeholder-fill.ts';

/** The placeholder types that never show the chrome: the footer row. */
const FOOTERS = new Set(['dt', 'ftr', 'sldNum', 'hdr']);
/** An insert button: a 28 px square, 8 px apart. */
const BUTTON_PX = 28;
const GAP_PX = 8;
/** A new table from a placeholder button: 3 columns and 3 rows, as a start. */
const TABLE_SIZE = 3;

/** What the placeholder chrome needs from the surface. */
export interface PlaceholderChromeHost {
  readonly selection: () => EditorSelection;
  readonly readOnly: () => boolean;
  readonly label: (key: string) => string;
  select(selection: EditorSelection): void;
  run(command: EditorCommand): void;
  /** Ask the user for a picture file and insert it; without it, no picture button shows. */
  readonly pickImage?: () => void;
}

type Insert = 'table' | 'picture';

/** An empty placeholder of the slide, top level and not turned: the chrome draws on its box. */
interface Empty {
  readonly item: Extract<LayoutItem, { kind: 'shape' }>;
  readonly type: string;
}

function emptyPlaceholders(layout: SlideLayout): Empty[] {
  const out: Empty[] = [];
  for (const item of layout.items) {
    if (item.kind !== 'shape') continue;
    const el = item.element;
    const type = el.placeholder?.type ?? (el.placeholder ? 'obj' : undefined);
    if (!type || el.layer !== 'slide' || el.hidden || FOOTERS.has(type)) continue;
    if (el.xfrm.rot !== 0 || el.xfrm.cx <= 0 || el.xfrm.cy <= 0) continue;
    const text = el.text?.paragraphs.some((p) =>
      p.runs.some((r) => r.kind === 'field' || r.text.length > 0)
    );
    if (!text) out.push({ item, type });
  }
  return out;
}

export class PlaceholderChrome {
  constructor(private readonly host: PlaceholderChromeHost) {}

  /** The outlines and buttons for `layout`; `px` is one screen pixel in slide units. */
  elements(doc: Document, layout: SlideLayout, px: number): Element[] {
    if (this.host.readOnly()) return [];
    const sel = this.host.selection();
    const editing = sel.kind === 'text' ? sel.nodeId : null;
    const parts: Element[] = [];
    for (const { item, type } of emptyPlaceholders(layout)) {
      const id = item.element.nodeId;
      // The placeholder being typed in shows the caret and its own outline instead.
      if (id === editing) continue;
      const { x, y, cx, cy } = item.element.xfrm;
      const box = doc.createElementNS(SVG_NS, 'rect');
      box.setAttribute('class', 'pptx-surface__placeholder-outline');
      box.setAttribute('x', String(x));
      box.setAttribute('y', String(y));
      box.setAttribute('width', String(cx));
      box.setAttribute('height', String(cy));
      box.setAttribute('stroke-width', String(px));
      box.setAttribute('stroke-dasharray', `${px} ${2 * px}`);
      parts.push(box);
      const inserts: Insert[] = [];
      if (TABLE_PLACEHOLDERS.has(type)) inserts.push('table');
      if (PICTURE_PLACEHOLDERS.has(type) && this.host.pickImage) inserts.push('picture');
      const width = (inserts.length * BUTTON_PX + (inserts.length - 1) * GAP_PX) * px;
      // A box too small for the buttons shows the outline only.
      if (!inserts.length || width > cx * 0.8 || BUTTON_PX * px > cy * 0.6) continue;
      let left = x + (cx - width) / 2;
      const top = y + cy / 2 - (BUTTON_PX * px) / 2;
      for (const insert of inserts) {
        parts.push(this.#button(doc, id, insert, left, top, px));
        left += (BUTTON_PX + GAP_PX) * px;
      }
    }
    return parts;
  }

  #button(doc: Document, id: string, insert: Insert, x: number, y: number, px: number): Element {
    const size = BUTTON_PX * px;
    const g = doc.createElementNS(SVG_NS, 'g');
    g.setAttribute('class', 'pptx-surface__placeholder-insert');
    g.setAttribute('data-placeholder-insert', `${id}|${insert}`);
    // Pointer chrome: keyboard users insert from the Insert menu, which fills the placeholder too.
    g.setAttribute('aria-hidden', 'true');
    g.setAttribute('transform', `translate(${x} ${y})`);
    const title = doc.createElementNS(SVG_NS, 'title');
    title.textContent = this.host.label(
      insert === 'table' ? 'canvas.placeholderTable' : 'canvas.placeholderPicture'
    );
    const box = doc.createElementNS(SVG_NS, 'rect');
    box.setAttribute('class', 'pptx-surface__placeholder-insert-box');
    box.setAttribute('width', String(size));
    box.setAttribute('height', String(size));
    box.setAttribute('rx', String(4 * px));
    box.setAttribute('stroke-width', String(px));
    // The 960-unit symbol at 20 px, centered in the box.
    const icon = doc.createElementNS(SVG_NS, 'g');
    const k = (20 * px) / 960;
    icon.setAttribute('transform', `translate(${4 * px} ${24 * px}) scale(${k})`);
    icon.setAttribute('class', 'pptx-surface__placeholder-insert-icon');
    for (const d of ICON_PATHS[insert === 'table' ? 'table' : 'image'] ?? []) {
      const path = doc.createElementNS(SVG_NS, 'path');
      path.setAttribute('d', d);
      icon.append(path);
    }
    g.append(title, box, icon);
    return g;
  }

  /**
   * A press on `target`: an insert button selects its placeholder and inserts a table or asks
   * for a picture, which then fills it. Returns whether the press was on a button.
   */
  press(layout: SlideLayout, target: Element): boolean {
    if (this.host.readOnly()) return false;
    const hit = target
      .closest?.('[data-placeholder-insert]')
      ?.getAttribute('data-placeholder-insert');
    if (!hit) return false;
    const [id, insert] = hit.split('|');
    // A peer can fill or delete the placeholder between the paint and the press.
    if (!emptyPlaceholders(layout).some((e) => e.item.element.nodeId === id)) return true;
    this.host.select({ kind: 'shapes', nodeIds: [id!] });
    if (insert === 'table')
      this.host.run({ type: 'insertTable', rows: TABLE_SIZE, cols: TABLE_SIZE });
    else this.host.pickImage?.();
    return true;
  }
}
