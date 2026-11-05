import { describe, expect, test } from 'bun:test';
import { buildPptx, shape } from '../../../../../scripts/fixtures/builder.ts';
import { path, readPptx, relationshipsOf, serializeXmlPart, xmlPartOf } from '../../store/index.ts';
import {
  createServerAutomationHost,
  type AutomationBatchResponse,
  type AutomationHandle,
  type AutomationHost,
  type AutomationOperation,
  type AutomationValue,
} from '../index.ts';

const IN = 914_400;
/** A 1 by 1 PNG. */
const PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';

function deck(background?: string): Uint8Array {
  return buildPptx({
    slides: [
      {
        layout: 'blank',
        ...(background ? { background } : {}),
        shapes: shape({ id: 2, name: 'Box', x: IN, y: IN, cx: IN, cy: IN }),
      },
      { layout: 'blank', shapes: '' },
    ],
  });
}

function open(bytes = deck()): AutomationHost {
  const opened = createServerAutomationHost(bytes);
  if (!opened.ok) throw new Error(JSON.stringify(opened.error));
  return opened.host;
}

function exec(host: AutomationHost, operations: AutomationOperation[]): AutomationBatchResponse {
  return host.execute({ operations });
}

function ok(host: AutomationHost, operations: AutomationOperation[]): AutomationValue[] {
  const response = exec(host, operations);
  if (!response.ok) throw new Error(JSON.stringify(response.results));
  return response.results.map((r) => (r.status === 'ok' ? r.value : { kind: 'null' as const }));
}

function handle(value: AutomationValue | undefined): AutomationHandle {
  if (value?.kind !== 'handle') throw new Error(`not a handle: ${JSON.stringify(value)}`);
  return value.handle;
}

/** The first slide, its layout, and its master. */
function owners(host: AutomationHost): {
  slide: AutomationHandle;
  layout: AutomationHandle;
  master: AutomationHandle;
} {
  const [slide, layout, master] = ok(host, [
    { op: 'getSlide', key: { index: 0 } },
    { op: 'getSlideLayout', slide: { result: 0 } },
    { op: 'getSlideMaster', owner: { result: 0 } },
  ]);
  return { slide: handle(slide), layout: handle(layout), master: handle(master) };
}

function background(host: AutomationHost, owner: AutomationHandle) {
  const [value] = ok(host, [{ op: 'readBackground', owner }]);
  if (value?.kind !== 'background') throw new Error(JSON.stringify(value));
  return value.background;
}

/** Save, check that the bytes read back, and open them again. */
function reopen(host: AutomationHost): { host: AutomationHost; bytes: Uint8Array } {
  const saved = host.save();
  if (!saved.ok) throw new Error(saved.error.code);
  expect(readPptx(saved.bytes).ok).toBe(true);
  return { host: open(saved.bytes), bytes: saved.bytes };
}

/** The `p:bg` XML of a part in saved bytes, or an empty string. */
function bgXml(bytes: Uint8Array, part: string): string {
  const read = readPptx(bytes);
  if (!read.ok) throw new Error('the saved presentation does not read back');
  const xml = xmlPartOf(read.pkg, part);
  if (!xml) throw new Error(`no part ${part}`);
  return path(xml.root, 'p:cSld', 'p:bg')
    ? serializeXmlPart(xml).replace(/.*?(<p:bg>.*<\/p:bg>).*/s, '$1')
    : '';
}

const SLIDE1 = '/ppt/slides/slide1.xml';

describe('automation host: backgrounds', () => {
  test('a slide shows the master background until it has its own', () => {
    const host = open();
    const { slide, layout, master } = owners(host);
    for (const owner of [slide, layout])
      expect(background(host, owner)).toEqual({
        followsMaster: true,
        graphicsHidden: false,
        fill: {
          type: 'solid',
          color: '#FFFFFF',
          transparency: 0,
          gradientType: null,
          pattern: null,
          foregroundColor: null,
          backgroundColor: null,
        },
      });
    // A master has no level above it.
    expect(background(host, master)).toMatchObject({
      followsMaster: false,
      fill: { type: 'solid' },
    });
  });

  test('a solid fill gives the slide its own background and survives a save', () => {
    const host = open();
    const { slide } = owners(host);
    ok(host, [
      {
        op: 'setBackground',
        owner: slide,
        background: { kind: 'solid', color: '#FF0000', transparency: 0.25 },
      },
    ]);
    expect(background(host, slide)).toMatchObject({
      followsMaster: false,
      fill: { type: 'solid', color: '#FF0000', transparency: 0.25 },
    });
    const again = reopen(host);
    expect(bgXml(again.bytes, SLIDE1)).toContain('<a:srgbClr val="FF0000"><a:alpha val="75000"/>');
    expect(background(again.host, owners(again.host).slide)).toMatchObject({
      fill: { color: '#FF0000', transparency: 0.25 },
    });
    // The second slide still follows the master.
    const [other] = ok(again.host, [{ op: 'getSlide', key: { index: 1 } }]);
    expect(background(again.host, handle(other)).followsMaster).toBe(true);
  });

  test('a transparency write keeps an own theme color', () => {
    const host = open(
      deck(
        '<p:bg><p:bgPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'
      )
    );
    const { slide } = owners(host);
    ok(host, [
      { op: 'setBackground', owner: slide, background: { kind: 'solid', transparency: 0.5 } },
    ]);
    expect(background(host, slide).fill).toMatchObject({ color: '#4472C4', transparency: 0.5 });
    const xml = bgXml(reopen(host).bytes, SLIDE1);
    expect(xml).toContain('<a:schemeClr val="accent1"><a:alpha val="50000"/></a:schemeClr>');
    expect(xml).toContain('<a:effectLst/>');
  });

  test('gradient writes keep the stops and change the shading', () => {
    const host = open();
    const { slide } = owners(host);
    ok(host, [{ op: 'setBackground', owner: slide, background: { kind: 'gradient' } }]);
    expect(background(host, slide).fill).toMatchObject({
      type: 'gradient',
      gradientType: 'linear',
    });
    for (const gradientType of ['circle', 'rect', 'shape', 'linear'] as const) {
      ok(host, [
        { op: 'setBackground', owner: slide, background: { kind: 'gradient', gradientType } },
      ]);
      expect(background(host, slide).fill.gradientType).toBe(gradientType);
    }
    const xml = bgXml(reopen(host).bytes, SLIDE1);
    expect(xml).toContain('<a:schemeClr val="bg1"/>');
    expect(xml).toContain('<a:lin ang="5400000" scaled="0"/>');
    const refused = exec(host, [
      {
        op: 'setBackground',
        owner: slide,
        background: { kind: 'gradient', gradientType: 'spiral' as 'linear' },
      },
    ]);
    expect(refused.results[0]).toMatchObject({
      error: { code: 'invalid-argument', detail: 'gradientType' },
    });
  });

  test('a title shade reads as shadeToTitle, and a write removes it', () => {
    const host = open(
      deck(
        '<p:bg><p:bgPr shadeToTitle="1"><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="000000"/></a:gs>' +
          '<a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst><a:path path="rect"/></a:gradFill></p:bgPr></p:bg>'
      )
    );
    const { slide } = owners(host);
    expect(background(host, slide).fill.gradientType).toBe('shadeToTitle');
    ok(host, [
      {
        op: 'setBackground',
        owner: slide,
        background: { kind: 'gradient', gradientType: 'circle' },
      },
    ]);
    expect(background(host, slide).fill.gradientType).toBe('circle');
    const xml = bgXml(reopen(host).bytes, SLIDE1);
    expect(xml).not.toContain('shadeToTitle');
    expect(xml).toContain('<a:srgbClr val="000000"/>');
  });

  test('pattern writes set the preset and colors, and keep what they leave out', () => {
    const host = open();
    const { slide } = owners(host);
    ok(host, [
      {
        op: 'setBackground',
        owner: slide,
        background: {
          kind: 'pattern',
          pattern: 'dkDnDiag',
          foregroundColor: '#112233',
          backgroundColor: '#445566',
        },
      },
    ]);
    ok(host, [
      {
        op: 'setBackground',
        owner: slide,
        background: { kind: 'pattern', foregroundColor: '#ABCDEF' },
      },
    ]);
    expect(background(host, slide).fill).toMatchObject({
      type: 'pattern',
      pattern: 'dkDnDiag',
      foregroundColor: '#ABCDEF',
      backgroundColor: '#445566',
    });
    reopen(host);
    for (const bad of [{ pattern: 'plaidish' }, { foregroundColor: 'url(x)' }]) {
      const response = exec(host, [
        { op: 'setBackground', owner: slide, background: { kind: 'pattern', ...bad } },
      ]);
      expect(response.results[0]).toMatchObject({ error: { code: 'invalid-argument' } });
    }
  });

  test('a picture fill adds a media part, and a transparency write keeps the picture', () => {
    const host = open();
    const { slide } = owners(host);
    const refused = exec(host, [
      { op: 'setBackground', owner: slide, background: { kind: 'picture' } },
    ]);
    expect(refused.results[0]).toMatchObject({
      error: { code: 'invalid-argument', detail: 'imageBase64' },
    });
    ok(host, [{ op: 'setBackground', owner: slide, background: { kind: 'picture', base64: PNG } }]);
    expect(background(host, slide).fill).toMatchObject({ type: 'picture', transparency: 0 });
    ok(host, [
      { op: 'setBackground', owner: slide, background: { kind: 'picture', transparency: 0.4 } },
    ]);
    expect(background(host, slide).fill).toMatchObject({ type: 'picture', transparency: 0.4 });
    const again = reopen(host);
    expect(bgXml(again.bytes, SLIDE1)).toContain('<a:alphaModFix amt="60000"/>');
    expect(background(again.host, owners(again.host).slide).fill.type).toBe('picture');
    for (const base64 of ['', 'not base64!', 'R0lG']) {
      const response = exec(host, [
        { op: 'setBackground', owner: slide, background: { kind: 'picture', base64 } },
      ]);
      expect(response.results[0]).toMatchObject({ error: { code: 'invalid-argument' } });
    }
  });

  test('a transparency write keeps the crop and the tile of an own picture', () => {
    const bytes = buildPptx({
      media: [
        {
          name: 'bg.png',
          bytes: Uint8Array.from(atob(PNG), (c) => c.charCodeAt(0)),
          contentType: 'image/png',
        },
      ],
      slides: [
        {
          layout: 'blank',
          rels: [
            [
              'rId9',
              'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
              '../media/bg.png',
            ],
          ],
          background:
            '<p:bg><p:bgPr><a:blipFill><a:blip r:embed="rId9"><a:alphaModFix amt="50000"/><a:grayscl/></a:blip>' +
            '<a:srcRect l="1000"/><a:tile tx="0" ty="0" sx="50000" sy="50000" flip="none" algn="tl"/></a:blipFill>' +
            '<a:effectLst/></p:bgPr></p:bg>',
          shapes: '',
        },
      ],
    });
    const host = open(bytes);
    const { slide } = owners(host);
    expect(background(host, slide).fill).toMatchObject({ type: 'picture', transparency: 0.5 });
    ok(host, [
      { op: 'setBackground', owner: slide, background: { kind: 'picture', transparency: 0.25 } },
    ]);
    const xml = bgXml(reopen(host).bytes, SLIDE1);
    expect(xml).toContain(
      '<a:blip r:embed="rId9"><a:alphaModFix amt="75000"/><a:grayscl/></a:blip>'
    );
    expect(xml).toContain('<a:srcRect l="1000"/><a:tile');
  });

  test('own copies the shown background with its picture, and follow removes it', () => {
    const host = open();
    const { slide, layout } = owners(host);
    ok(host, [
      { op: 'setBackground', owner: layout, background: { kind: 'picture', base64: PNG } },
    ]);
    expect(background(host, slide)).toMatchObject({
      followsMaster: true,
      fill: { type: 'picture' },
    });
    ok(host, [{ op: 'setBackground', owner: slide, background: { kind: 'own' } }]);
    expect(background(host, slide)).toMatchObject({
      followsMaster: false,
      fill: { type: 'picture' },
    });
    const again = reopen(host);
    const read = readPptx(again.bytes);
    if (!read.ok) throw new Error('no read');
    const embed = /r:embed="([^"]+)"/.exec(bgXml(again.bytes, SLIDE1))?.[1];
    const rel = relationshipsOf(read.pkg, SLIDE1).find((r) => r.id === embed);
    expect(rel?.target).toMatch(/^\/ppt\/media\/image/);
    ok(host, [{ op: 'setBackground', owner: slide, background: { kind: 'follow' } }]);
    expect(background(host, slide).followsMaster).toBe(true);
    expect(bgXml(reopen(host).bytes, SLIDE1)).toBe('');
  });

  test('own on a slide with no background anywhere writes white', () => {
    const host = open();
    const { slide, master } = owners(host);
    // A master background with no fill leaves the default white.
    ok(host, [
      { op: 'setBackground', owner: master, background: { kind: 'solid', color: '#00FF00' } },
    ]);
    expect(background(host, slide).fill.color).toBe('#00FF00');
    ok(host, [{ op: 'setBackground', owner: slide, background: { kind: 'own' } }]);
    expect(bgXml(reopen(host).bytes, SLIDE1)).toContain('00FF00');
  });

  test('showMasterSp hides the graphics of the level above on a slide and a layout', () => {
    const host = open();
    const { slide, layout, master } = owners(host);
    ok(host, [
      { op: 'setBackground', owner: slide, background: { kind: 'graphicsHidden', hidden: true } },
      { op: 'setBackground', owner: layout, background: { kind: 'graphicsHidden', hidden: true } },
    ]);
    expect(background(host, slide).graphicsHidden).toBe(true);
    expect(background(host, layout).graphicsHidden).toBe(true);
    const again = reopen(host).host;
    expect(background(again, owners(again).slide).graphicsHidden).toBe(true);
    ok(host, [
      { op: 'setBackground', owner: slide, background: { kind: 'graphicsHidden', hidden: false } },
    ]);
    expect(background(host, slide).graphicsHidden).toBe(false);
    for (const write of [
      { kind: 'follow' },
      { kind: 'own' },
      { kind: 'graphicsHidden', hidden: true },
    ] as const) {
      const response = exec(host, [{ op: 'setBackground', owner: master, background: write }]);
      expect(response.results[0]).toMatchObject({ error: { code: 'invalid-argument' } });
    }
  });

  test('a layout background shows on the slides that follow it', () => {
    const host = open();
    const { slide, layout } = owners(host);
    ok(host, [
      { op: 'setBackground', owner: layout, background: { kind: 'solid', color: 'blue' } },
    ]);
    expect(background(host, layout)).toMatchObject({
      followsMaster: false,
      fill: { color: '#0000FF' },
    });
    expect(background(host, slide)).toMatchObject({
      followsMaster: true,
      fill: { color: '#0000FF' },
    });
  });

  test('refuses bad input and owners that have no background', () => {
    const host = open();
    const { slide } = owners(host);
    const [box] = ok(host, [{ op: 'getShape', container: slide, key: { index: 0 } }]);
    const cases: AutomationOperation[] = [
      { op: 'setBackground', owner: slide, background: { kind: 'solid', transparency: 2 } },
      { op: 'setBackground', owner: slide, background: { kind: 'solid', color: 'url(x)' } },
      {
        op: 'setBackground',
        owner: slide,
        background: { kind: 'graphicsHidden', hidden: 'yes' as unknown as boolean },
      },
      { op: 'setBackground', owner: slide, background: { kind: 'other' } as never },
      { op: 'setBackground', owner: slide, background: null as never },
    ];
    for (const op of cases)
      expect(exec(host, [op]).results[0]).toMatchObject({ error: { code: 'invalid-argument' } });
    expect(exec(host, [{ op: 'readBackground', owner: handle(box) }]).results[0]).toMatchObject({
      error: { code: 'invalid-handle' },
    });
  });
});

describe('automation host: theme colors', () => {
  test('reads the theme of the master from a slide, a layout, and a master', () => {
    const host = open();
    const { slide, layout, master } = owners(host);
    for (const owner of [slide, layout, master])
      expect(ok(host, [{ op: 'readThemeColor', owner, slot: 'accent1' }])[0]).toEqual({
        kind: 'text',
        text: '#4472C4',
      });
    // `a:sysClr` reads its last color.
    expect(ok(host, [{ op: 'readThemeColor', owner: slide, slot: 'lt1' }])[0]).toEqual({
      kind: 'text',
      text: '#FFFFFF',
    });
  });

  test('a write changes the theme part for every owner, and survives a save', () => {
    const host = open();
    const { slide, master } = owners(host);
    ok(host, [{ op: 'setThemeColor', owner: slide, slot: 'accent2', color: '#123456' }]);
    expect(ok(host, [{ op: 'readThemeColor', owner: master, slot: 'accent2' }])[0]).toEqual({
      kind: 'text',
      text: '#123456',
    });
    const again = reopen(host).host;
    expect(
      ok(again, [{ op: 'readThemeColor', owner: owners(again).layout, slot: 'accent2' }])[0]
    ).toEqual({
      kind: 'text',
      text: '#123456',
    });
    // A background in a theme color follows the new value.
    ok(host, [{ op: 'setThemeColor', owner: slide, slot: 'lt1', color: '#FAFAFA' }]);
    expect(background(host, slide).fill.color).toBe('#FAFAFA');
  });

  test('refuses unknown slots, bad colors, and owners that have no theme', () => {
    const host = open();
    const { slide } = owners(host);
    for (const op of [
      { op: 'readThemeColor', owner: slide, slot: 'bg1' },
      { op: 'readThemeColor', owner: slide, slot: '__proto__' },
      { op: 'setThemeColor', owner: slide, slot: 'accent1', color: 'url(x)' },
      { op: 'setThemeColor', owner: slide, slot: 'none', color: '#000000' },
    ] as AutomationOperation[])
      expect(exec(host, [op]).results[0]).toMatchObject({ error: { code: 'invalid-argument' } });
    const [presentation] = ok(host, [{ op: 'getPresentation' }]);
    expect(
      exec(host, [{ op: 'readThemeColor', owner: handle(presentation), slot: 'accent1' }])
        .results[0]
    ).toMatchObject({ error: { code: 'invalid-handle' } });
  });
});
