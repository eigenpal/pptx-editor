// The SVG paint builds an effects renderer's plain data (`EffectNode`) with DOM APIs, and only
// effect elements and attributes: a renderer can never add script or fetch a resource.

import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { effectElement } from '../paint-effects.ts';

const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

describe('effect nodes', () => {
  test('build as SVG elements with their attributes and children', () => {
    const el = effectElement(doc, {
      tag: 'filter',
      attrs: { id: 'f1', x: '0' },
      children: [{ tag: 'feGaussianBlur', attrs: { stdDeviation: '2' } }],
    });
    expect(el.namespaceURI).toBe('http://www.w3.org/2000/svg');
    expect(el.getAttribute('id')).toBe('f1');
    expect(el.firstElementChild?.getAttribute('stdDeviation')).toBe('2');
  });

  test('take attribute values as text, never as markup', () => {
    const el = effectElement(doc, { tag: 'g', attrs: { 'data-effect': '"><script>x</script>' } });
    expect(el.getAttribute('data-effect')).toBe('"><script>x</script>');
    expect(el.children.length).toBe(0);
  });

  test('build no script, foreign content, or link', () => {
    for (const tag of ['script', 'foreignObject', 'a', 'image', 'style']) {
      const el = effectElement(doc, { tag, attrs: { href: '#x' }, children: [] });
      expect(el.localName).toBe('g');
      expect(el.attributes.length).toBe(0);
    }
  });

  test('drop event handlers and references outside the slide', () => {
    const el = effectElement(doc, {
      tag: 'use',
      attrs: { href: 'https://example.com/a.svg#x', onload: 'x()', OnClick: 'x()' },
    });
    expect(el.attributes.length).toBe(0);
    const rect = effectElement(doc, {
      tag: 'rect',
      attrs: { fill: 'url(https://example.com/p.svg#g)', mask: "url('#m1')", filter: 'url(#f1)' },
    });
    expect(rect.getAttribute('fill')).toBeNull();
    expect(rect.getAttribute('mask')).toBe("url('#m1')");
    expect(rect.getAttribute('filter')).toBe('url(#f1)');
    expect(effectElement(doc, { tag: 'use', attrs: { href: '#c1' } }).getAttribute('href')).toBe(
      '#c1'
    );
  });
});
