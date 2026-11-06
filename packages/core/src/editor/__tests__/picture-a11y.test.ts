// Pictures for assistive technology on the editable slide: a picture with alt text is an image
// named by its description, a decorative picture or one with no description is hidden, the
// selection overlay is hidden, and the live region says the kind, the name, and the alt text of
// the shape that Tab selects. The slide show exposes the same text and alt text.

import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { createTranslator, isTranslationKey } from '@pptx-editor.dev/i18n';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { createPptxEditor, paintSlideshowSlide, type PptxEditorInstance } from '../index.ts';
import { nonVisualOf } from '../picture-commands.ts';
import { nodeIndex, xmlPartOf, type Element as XmlElement } from '../../store/index.ts';

const IN = 914_400;
const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const DECORATIVE =
  '<a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}">' +
  '<adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/>' +
  '</a:ext></a:extLst>';

function pic(id: number, name: string, x: number, cNvPr = '', inner = ''): string {
  return (
    `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${name}"${cNvPr}>${inner}</p:cNvPr>` +
    '<p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
    '<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
    `<p:spPr><a:xfrm><a:off x="${x}" y="${IN}"/><a:ext cx="${IN}" cy="${IN}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
  );
}

/** A picture with a description, a decorative one, one with none, a shape, and a default name. */
function deck(): Uint8Array {
  return buildPptx({
    slides: [
      {
        layout: 'blank',
        rels: [['rIdImg', R_IMAGE, '../media/image1.png']],
        shapes:
          pic(2, 'Landscape', IN, ' descr="A dusk sky over a lake"') +
          pic(3, 'Ornament', 3 * IN, ' descr="Swirl"', DECORATIVE) +
          pic(4, 'Photo', 5 * IN) +
          shape({ id: 5, name: 'Box', x: 7 * IN, y: IN, cx: IN, cy: IN, text: para(run('Trip')) }) +
          pic(6, 'Picture 6', 9 * IN, ' descr="A bar chart"') +
          shape({ id: 7, name: 'Shapely', x: 11 * IN, y: IN, cx: IN, cy: IN }),
      },
    ],
    media: [
      { name: 'image1.png', bytes: new Uint8Array([137, 80, 78, 71]), contentType: 'image/png' },
    ],
  });
}

// A DOM for this file only: the engine test run registers no global DOM.
const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

async function attached() {
  const g = globalThis as { ResizeObserver?: unknown };
  g.ResizeObserver ??= (win as unknown as { ResizeObserver: unknown }).ResizeObserver;
  const t = createTranslator();
  const e = createPptxEditor({
    fixedMeasurer: true,
    substitutes: false,
    translate: (key, params) => (isTranslationKey(key) ? t(key, params) : key),
  });
  expect((await e.load(deck())).ok).toBe(true);
  const host = doc.createElement('div') as unknown as HTMLElement;
  doc.body.append(host as unknown as Node);
  e.attach(host);
  const root = host.querySelector('.pptx-surface') as HTMLElement;
  const tab = (): void => {
    root.dispatchEvent(
      new win.KeyboardEvent('keydown', {
        key: 'Tab',
        bubbles: true,
        cancelable: true,
      }) as unknown as Event
    );
  };
  return { e, host, root, tab };
}

const picture = (host: HTMLElement, e: PptxEditorInstance, name: string): Element => {
  const item = e.session!.layout(0)!.items.find((i) => i.element.name === name)!;
  return host.querySelector(`.pptx-surface__slide [data-node-id="${item.element.nodeId}"]`)!;
};

describe('pictures for assistive technology', () => {
  test('a described picture is a named image; decorative and undescribed ones are hidden', async () => {
    const { e, host } = await attached();
    const landscape = picture(host, e, 'Landscape');
    expect(landscape.getAttribute('role')).toBe('img');
    expect(landscape.getAttribute('aria-label')).toBe('A dusk sky over a lake');
    for (const name of ['Ornament', 'Photo']) {
      expect(picture(host, e, name).getAttribute('aria-hidden')).toBe('true');
      expect(picture(host, e, name).getAttribute('role')).toBeNull();
    }
    expect(picture(host, e, 'Box').getAttribute('aria-hidden')).toBeNull();
    // The selection chrome is drawing only: the live region says what is selected.
    expect(host.querySelector('.pptx-surface__overlay')?.getAttribute('aria-hidden')).toBe('true');
    e.destroy();
  });

  test('Tab says the kind, the name, and the alt text of the selected shape', async () => {
    const { e, root, tab } = await attached();
    const said = () => root.querySelector('[aria-live]')?.textContent;
    tab();
    expect(said()).toBe('Picture Landscape, A dusk sky over a lake, selected');
    tab();
    // A decorative picture keeps no description in the announcement.
    expect(said()).toBe('Picture Ornament, selected');
    tab();
    expect(said()).toBe('Picture Photo, selected');
    tab();
    expect(said()).toBe('Shape Box, selected');
    tab();
    // A name that starts with the kind says the kind once.
    expect(said()).toBe('Picture 6, A bar chart, selected');
    tab();
    // The kind is a whole word: "Shapely" does not start with the kind "Shape".
    expect(said()).toBe('Shape Shapely, selected');
    e.destroy();
  });

  test('new alt text names the picture and the announcement at once', async () => {
    const { e, host, root, tab } = await attached();
    tab();
    tab();
    tab();
    expect(e.exec({ type: 'setAltText', description: 'A portrait' }).ok).toBe(true);
    expect(picture(host, e, 'Photo').getAttribute('role')).toBe('img');
    expect(picture(host, e, 'Photo').getAttribute('aria-label')).toBe('A portrait');
    expect(root.querySelector('[aria-live]')?.textContent).toBe(
      'Picture Photo, A portrait, selected'
    );
    expect(e.exec({ type: 'setAltText', decorative: true }).ok).toBe(true);
    expect(picture(host, e, 'Photo').getAttribute('aria-hidden')).toBe('true');
    expect(root.querySelector('[aria-live]')?.textContent).toBe('Picture Photo, selected');
    e.destroy();
  });

  test('undo, redo, and a remote edit re-derive the name and the announcement', async () => {
    const { e, host, root, tab } = await attached();
    const said = () => root.querySelector('[aria-live]')?.textContent;
    tab();
    tab();
    tab();
    expect(e.exec({ type: 'setAltText', description: 'A portrait' }).ok).toBe(true);
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(picture(host, e, 'Photo').getAttribute('aria-hidden')).toBe('true');
    expect(said()).toBe('Picture Photo, selected');
    expect(e.exec({ type: 'redo' }).ok).toBe(true);
    expect(picture(host, e, 'Photo').getAttribute('aria-label')).toBe('A portrait');
    // A peer's edit arrives as a remote transaction: the paint derives the ARIA again.
    const item = e.session!.layout(0)!.items.find((i) => i.element.name === 'Landscape')!;
    const part = item.element.partName;
    const pic = nodeIndex(xmlPartOf(e.session!.store.package, part)!).get(item.element.nodeId)!;
    const cNvPr = nonVisualOf(pic.node as XmlElement)!;
    const r = e.session!.store.transact(
      [
        {
          op: 'setAttribute',
          part,
          node: cNvPr.id,
          name: { ns: '', local: 'descr' },
          value: 'A storm over a lake',
        },
      ],
      { origin: 'remote' }
    );
    expect(r.ok).toBe(true);
    expect(picture(host, e, 'Landscape').getAttribute('aria-label')).toBe('A storm over a lake');
    e.destroy();
  });

  test('the slide show reads the slide text and the picture alt text, not one image', async () => {
    const { e } = await attached();
    const svg = paintSlideshowSlide(doc, e.session!, 0)!;
    expect(svg.getAttribute('role')).toBe('group');
    expect(svg.textContent).toContain('Trip');
    const named = [...svg.querySelectorAll('[role="img"]')].map((g) =>
      g.getAttribute('aria-label')
    );
    expect(named).toEqual(['A dusk sky over a lake', 'A bar chart']);
    expect(svg.querySelectorAll('[aria-hidden="true"]').length).toBe(2);
    e.destroy();
  });
});
