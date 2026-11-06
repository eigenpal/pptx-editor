/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Slide, layout, and master backgrounds and the theme color scheme, end to end on a server
// runtime, with a save and reopen after the edits.

import { describe, expect, test } from 'bun:test';
import {
  isPptxEditorError,
  SlideBackgroundFillType,
  SlideBackgroundPatternFillType,
  ThemeColor,
  type RequestContext,
} from '../index.ts';
import { open, refusedAtSync, rejection, reopen } from './support.ts';

const code = (error: unknown): string | undefined =>
  isPptxEditorError(error) ? error.code : String(error);

/** A 1 by 1 PNG. */
const PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';

const first = (context: RequestContext) => context.presentation.slides.getItemAt(0);

describe('slide backgrounds', () => {
  test('a slide reads the master background it follows', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const background = first(context).background.load();
      const fill = background.fill.load('type');
      const solid = background.fill.getSolidFillOrNullObject().load();
      const gradient = background.fill.getGradientFillOrNullObject();
      await context.sync();
      expect(background.isMasterBackgroundFollowed).toBe(true);
      expect(background.areBackgroundGraphicsHidden).toBe(false);
      expect(fill.type).toBe(SlideBackgroundFillType.solid);
      expect([solid.isNullObject, solid.color, solid.transparency]).toEqual([false, '#FFFFFF', 0]);
      expect(gradient.isNullObject).toBe(true);
    });
  });

  test('a solid fill and its transparency survive a save', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      first(context).background.fill.setSolidFill({ color: '#336699', transparency: 0.2 });
      await context.sync();
      const solid = first(context).background.fill.getSolidFillOrNullObject();
      solid.transparency = 0.5;
      await context.sync();
    });
    const again = await reopen(runtime);
    await again.run(async (context) => {
      const background = first(context).background.load('isMasterBackgroundFollowed');
      const solid = background.fill.getSolidFillOrNullObject().load();
      await context.sync();
      expect(background.isMasterBackgroundFollowed).toBe(false);
      expect([solid.color, solid.transparency]).toEqual(['#336699', 0.5]);
      solid.color = 'red';
      await context.sync();
      solid.load('color');
      await context.sync();
      expect(solid.color).toBe('#FF0000');
    });
  });

  test('a gradient fill and its shading', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const fill = first(context).background.fill;
      fill.setGradientFill({ type: 'Radial' });
      await context.sync();
      const gradient = fill.getGradientFillOrNullObject().load('type');
      await context.sync();
      expect(gradient.type).toBe('Radial');
      gradient.type = 'Rectangular';
      await context.sync();
      gradient.load('type');
      fill.load('type');
      await context.sync();
      expect([fill.type, gradient.type]).toEqual(['Gradient', 'Rectangular']);
      expect(code(await refusedAtSync(context, () => (gradient.type = 'ShadeFromTitle')))).toBe(
        'NotImplemented'
      );
      expect(code(await refusedAtSync(context, () => (gradient.type = 'Unsupported')))).toBe(
        'InvalidArgument'
      );
    });
    const again = await reopen(runtime);
    await again.run(async (context) => {
      const gradient = first(context).background.fill.getGradientFillOrNullObject().load('type');
      await context.sync();
      expect(gradient.type).toBe('Rectangular');
    });
  });

  test('a pattern fill with colors, and every pattern writes', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const fill = first(context).background.fill;
      fill.setPatternFill({
        pattern: SlideBackgroundPatternFillType.diagonalBrick,
        foregroundColor: '#112233',
        backgroundColor: '#FFEEDD',
      });
      await context.sync();
      const pattern = fill.getPatternFillOrNullObject();
      pattern.foregroundColor = '#000080';
      await context.sync();
      pattern.load();
      await context.sync();
      expect(pattern.toJSON()).toEqual({
        backgroundColor: '#FFEEDD',
        foregroundColor: '#000080',
        pattern: 'DiagonalBrick',
      });
      for (const value of Object.values(SlideBackgroundPatternFillType)) {
        if (value === 'Unsupported') continue;
        pattern.pattern = value;
      }
      await context.sync();
      pattern.load('pattern');
      await context.sync();
      expect(pattern.pattern).toBe('ZigZag');
      expect(code(await refusedAtSync(context, () => (pattern.pattern = 'Unsupported')))).toBe(
        'InvalidArgument'
      );
    });
    await reopen(runtime);
  });

  test('a picture fill, its transparency, and a new image', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const fill = first(context).background.fill;
      fill.setPictureOrTextureFill({ imageBase64: PNG, transparency: 0.3 });
      await context.sync();
      const picture = fill.getPictureOrTextureFillOrNullObject().load('transparency');
      await context.sync();
      expect(picture.transparency).toBe(0.3);
      picture.setImage(PNG);
      picture.transparency = 0.6;
      await context.sync();
    });
    const again = await reopen(runtime);
    await again.run(async (context) => {
      const fill = first(context).background.fill.load('type');
      const picture = fill.getPictureOrTextureFillOrNullObject().load('transparency');
      await context.sync();
      expect([fill.type, picture.transparency]).toEqual(['PictureOrTexture', 0.6]);
    });
  });

  test('follow, own, reset, and hidden background graphics', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const background = first(context).background;
      background.isMasterBackgroundFollowed = false;
      background.areBackgroundGraphicsHidden = true;
      await context.sync();
      background.load();
      await context.sync();
      expect(background.toJSON()).toEqual({
        areBackgroundGraphicsHidden: true,
        isMasterBackgroundFollowed: false,
      });
      background.reset();
      await context.sync();
      background.load('isMasterBackgroundFollowed');
      await context.sync();
      expect(background.isMasterBackgroundFollowed).toBe(true);
      background.fill.setSolidFill({ color: '#00FF00' });
      background.isMasterBackgroundFollowed = true;
      await context.sync();
      background.load('isMasterBackgroundFollowed');
      await context.sync();
      expect(background.isMasterBackgroundFollowed).toBe(true);
    });
    const again = await reopen(runtime);
    await again.run(async (context) => {
      const background = first(context).background.load('areBackgroundGraphicsHidden');
      await context.sync();
      expect(background.areBackgroundGraphicsHidden).toBe(true);
    });
  });

  test('a layout and a master background show on the slides that follow them', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const slide = first(context);
      slide.slideMaster.background.fill.setSolidFill({ color: '#202020' });
      await context.sync();
      const fromMaster = slide.background.fill.getSolidFillOrNullObject().load('color');
      await context.sync();
      expect(fromMaster.color).toBe('#202020');
      const layout = slide.layout.background;
      layout.fill.setPatternFill({ pattern: 'Plaid' });
      layout.areBackgroundGraphicsHidden = true;
      await context.sync();
      layout.load();
      const fill = slide.background.fill.load('type');
      await context.sync();
      expect([layout.isMasterBackgroundFollowed, layout.areBackgroundGraphicsHidden]).toEqual([
        false,
        true,
      ]);
      expect(fill.type).toBe('Pattern');
      layout.reset();
      await context.sync();
      fill.load('type');
      await context.sync();
      expect(fill.type).toBe('Solid');
    });
    await reopen(runtime);
  });

  test('refuses bad options, values, and bytes at the sync', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const fill = first(context).background.fill;
      const cases: (() => unknown)[] = [
        () => fill.setSolidFill({ transparency: 2 }),
        () => fill.setSolidFill({ colour: '#000000' } as never),
        () => fill.setSolidFill('red' as never),
        () => fill.setPatternFill({ pattern: 'Stripes' as never }),
        () => fill.setGradientFill({ type: 'Spiral' as never }),
        () => fill.setPictureOrTextureFill({ transparency: -1 }),
        () => (first(context).background.areBackgroundGraphicsHidden = 'yes' as never),
      ];
      for (const call of cases)
        expect(code(await refusedAtSync(context, call))).toBe('InvalidArgument');
      expect(
        code(await refusedAtSync(context, () => fill.setGradientFill({ type: 'ShadeFromTitle' })))
      ).toBe('NotImplemented');
    });
    expect(
      code(
        await rejection(
          runtime.run(async (context) => {
            first(context).background.fill.setPictureOrTextureFill({
              imageBase64: 'not base64!',
              transparency: 0,
            });
            await context.sync();
          })
        )
      )
    ).toBe('InvalidArgument');
    expect(
      code(
        await rejection(
          runtime.run(async (context) => {
            // No picture to keep: the image is required.
            first(context).background.fill.setPictureOrTextureFill({ transparency: 0 });
            await context.sync();
          })
        )
      )
    ).toBe('InvalidArgument');
  });

  test('a write on a typed fill of another kind is a null object', async () => {
    const runtime = await open();
    const error = await rejection(
      runtime.run(async (context) => {
        const pattern = first(context).background.fill.getPatternFillOrNullObject();
        await context.sync();
        expect(pattern.isNullObject).toBe(true);
        pattern.pattern = 'Cross';
      })
    );
    expect(code(error)).toBe('InvalidObjectPath');
  });
});

describe('theme color scheme', () => {
  test('reads and writes theme colors through a slide, a layout, and a master', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const slide = first(context);
      const colors = [
        slide.themeColorScheme.getThemeColor(ThemeColor.accent1),
        slide.layout.themeColorScheme.getThemeColor('Light1'),
        slide.slideMaster.themeColorScheme.getThemeColor('Hyperlink'),
      ];
      await context.sync();
      expect(colors.map((c) => c.value)).toEqual(['#4472C4', '#FFFFFF', '#0563C1']);
      slide.themeColorScheme.setThemeColor('Accent1', '#AA5500');
      await context.sync();
    });
    const again = await reopen(runtime);
    await again.run(async (context) => {
      const color = first(context).slideMaster.themeColorScheme.getThemeColor('Accent1');
      await context.sync();
      expect(color.value).toBe('#AA5500');
      const scheme = first(context).themeColorScheme;
      expect(code(await refusedAtSync(context, () => scheme.getThemeColor('Mixed')))).toBe(
        'InvalidArgument'
      );
      expect(
        code(await refusedAtSync(context, () => scheme.setThemeColor('None', '#000000')))
      ).toBe('InvalidArgument');
      expect(
        code(await refusedAtSync(context, () => scheme.setThemeColor('Dark1', 7 as never)))
      ).toBe('InvalidArgument');
    });
    expect(
      code(
        await rejection(
          again.run(async (context) => {
            first(context).themeColorScheme.setThemeColor('Dark2', 'url(x)');
            await context.sync();
          })
        )
      )
    ).toBe('InvalidArgument');
  });
});
