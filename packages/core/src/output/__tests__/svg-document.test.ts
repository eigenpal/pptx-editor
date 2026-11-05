// The DOM-free SVG document: what `paintSlide` builds with when there is no DOM, and its
// serializer, which must escape every value.

import { describe, expect, test } from 'bun:test';
import { Window } from 'happy-dom';
import { buildPptx, para, run, shape, solid } from '../../../../../scripts/fixtures/builder.ts';
import { openPresentation } from '../../editor/index.ts';
import { paintSlide } from '../paint.ts';
import { SVG_NS } from '../svg.ts';
import { createSvgDocument, elementsWithAttribute, serializeSvgElement } from '../svg-document.ts';

describe('the DOM-free SVG document', () => {
  test('serializes elements, attributes, and text, with the SVG namespace on the root', () => {
    const doc = createSvgDocument();
    const root = doc.createElementNS(SVG_NS, 'svg');
    const text = doc.createElementNS(SVG_NS, 'text');
    text.setAttribute('x', '1');
    text.textContent = 'A & B';
    root.appendChild(text);
    root.insertBefore(doc.createElementNS(SVG_NS, 'defs'), root.firstChild);
    expect(serializeSvgElement(root)).toBe(
      `<svg xmlns="${SVG_NS}"><defs/><text x="1">A &amp; B</text></svg>`
    );
  });

  test('escapes markup in values and replaces characters that XML does not allow', () => {
    const doc = createSvgDocument();
    const root = doc.createElementNS(SVG_NS, 'svg');
    root.setAttribute('data-name', '"><script>x</script>\n');
    root.textContent = '</svg><svg onload="x">\u0001';
    const out = serializeSvgElement(root);
    expect(out).not.toContain('<script');
    expect(out).not.toContain('</svg><svg');
    expect(out).toContain('&quot;&gt;&lt;script&gt;');
    expect(out).toContain('�');
  });

  test('refuses a name that is not an XML name, and a cycle', () => {
    const doc = createSvgDocument();
    const root = doc.createElementNS(SVG_NS, 'svg');
    expect(() => root.setAttribute('a b', 'x')).toThrow();
    expect(() => doc.createElementNS(SVG_NS, 'g onload')).toThrow();
    const g = doc.createElementNS(SVG_NS, 'g');
    root.appendChild(g);
    expect(() => g.appendChild(root)).toThrow();
  });

  test('walks a tree by attribute, and moves a node that is inserted again', () => {
    const doc = createSvgDocument();
    const root = doc.createElementNS(SVG_NS, 'svg');
    const a = doc.createElementNS(SVG_NS, 'g');
    const b = doc.createElementNS(SVG_NS, 'rect');
    b.setAttribute('data-kind', 'background');
    a.appendChild(b);
    root.appendChild(a);
    expect(elementsWithAttribute(root, 'data-kind')).toEqual([b]);
    root.appendChild(b);
    expect(a.children.length).toBe(0);
    expect(b.parentNode).toBe(root);
  });

  test('a painted slide serializes to the same SVG as a DOM document gives', async () => {
    const opened = await openPresentation(
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes: shape({
              id: 2,
              name: 'Box',
              x: 914_400,
              y: 914_400,
              cx: 2_743_200,
              cy: 914_400,
              fill: solid('4472C4'),
              text: para(run('Tom & "Jerry" <b>')),
            }),
          },
        ],
      }),
      { substitutes: false }
    );
    if (!opened.ok) throw new Error('fixture');
    const layout = opened.session.layout(0)!;
    const options = { images: () => undefined, idPrefix: 'p', hitTargets: false };
    const window = new Window();
    const dom = paintSlide(window.document as unknown as Document, layout, options);
    const shim = paintSlide(createSvgDocument(), layout, options);
    // Parse both through the DOM, so only the markup, not its spelling, is compared.
    const parse = (svg: string): string =>
      new window.XMLSerializer().serializeToString(
        new window.DOMParser().parseFromString(svg, 'image/svg+xml').documentElement
      );
    const fromDom = new window.XMLSerializer().serializeToString(dom as never);
    const fromShim = serializeSvgElement(shim);
    expect(fromShim).toContain('&quot;Jerry&quot; &lt;b&gt;');
    expect(fromShim).toContain('data-kind="background"');
    expect(parse(fromShim)).toBe(parse(fromDom));
    opened.session.dispose();
    await window.happyDOM.close();
  });
});
