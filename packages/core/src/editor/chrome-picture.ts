// The `image` chrome group: the Picture Format tab (Adjust, Picture Styles,
// Accessibility, and Size), after DOCX Editor's `image` group (`image.altText`). The toolbar shows it while every selected element is a picture (`EditorSnapshot.image`). Kept apart from
// chrome-controls.ts, which spreads it into `CHROME_GROUPS` and routes its slots here.

import type { Editor, EditorCommand, ImageRecolor, ImageStyleId } from '../contracts/index.ts';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import type {
  ChromeControl,
  ChromeGroup,
  ChromeSlotId,
  ToolbarCommandState,
} from './chrome-controls.ts';
import { ICON_PATHS } from './icon-paths.ts';
import { PICTURE_STYLE_IDS, RECOLOR_PRESETS } from './picture-presets.ts';

const icon = (name: string): readonly string[] => ICON_PATHS[name] ?? [];

/**
 * The Brightness/Contrast presets of the Corrections gallery, in percent: a grid of
 * 25, contrast by row and brightness by column, each from -40 to 40 in steps of 20. The toolbar
 * draws them as that grid (`correctionGridRows`); a menu lists them in that order.
 */
const STEPS = [-40, -20, 0, 20, 40] as const;
const CORRECTIONS: readonly (readonly [number, number])[] = STEPS.flatMap((contrast) =>
  STEPS.map((brightness) => [brightness, contrast] as const)
);
/** The Transparency gallery, in percent. */
const TRANSPARENCY = [0, 15, 30, 50, 65, 80, 95] as const;
/** Crop: the Aspect Ratio menu (width:height), then Fill and Fit, in the menu's order. */
const CROP_FITS = [
  '1:1',
  '2:3',
  '3:4',
  '3:5',
  '4:5',
  '3:2',
  '4:3',
  '5:3',
  '5:4',
  '16:9',
  '16:10',
  'fill',
  'fit',
] as const;

const correctionValue = (b: number, c: number): string => `${b},${c}`;
/** A percent as the gallery labels it: `+20`, `0`, `-40`. */
const signed = (v: number): string => (v > 0 ? `+${v}` : String(v));

/**
 * The `image` group. Color, Transparency, Reset, Crop to shape, and Crop fit stay off the default
 * bar: Format > Picture and the picture context menu place them. @internal
 */
export const IMAGE_CHROME_GROUP: ChromeGroup = {
  id: 'image',
  labelKey: 'toolbar.groups.image',
  contextual: true,
  controls: [
    {
      id: 'corrections',
      shape: 'dropdown',
      labelKey: 'pictureFormat.corrections',
      toggle: 'choice',
      paths: icon('brightness_6'),
      options: CORRECTIONS.map(([b, c]) => ({
        value: correctionValue(b, c),
        labelKey: 'pictureFormat.correction',
        labelParams: { brightness: signed(b), contrast: signed(c) },
      })),
      state: { kind: 'value' },
    },
    {
      id: 'color',
      defaultToolbar: false,
      shape: 'dropdown',
      labelKey: 'pictureFormat.color',
      toggle: 'choice',
      paths: icon('palette'),
      options: RECOLOR_PRESETS.map((r) => ({
        value: r,
        labelKey: `pictureFormat.recolor.${r}` as TranslationKey,
      })),
      state: { kind: 'value' },
    },
    {
      id: 'transparency',
      defaultToolbar: false,
      shape: 'dropdown',
      labelKey: 'pictureFormat.transparency',
      toggle: 'choice',
      paths: icon('opacity'),
      options: TRANSPARENCY.map((t) => ({
        value: String(t),
        labelKey: 'pictureFormat.transparencyValue',
        labelParams: { value: t },
      })),
      state: { kind: 'value' },
    },
    {
      id: 'change',
      labelKey: 'pictureFormat.change',
      paths: icon('imagesmode'),
      state: { kind: 'replace' },
    },
    {
      id: 'reset',
      defaultToolbar: false,
      shape: 'dropdown',
      labelKey: 'pictureFormat.reset',
      paths: icon('restart_alt'),
      options: [
        { value: 'picture', labelKey: 'pictureFormat.resetImage' },
        { value: 'size', labelKey: 'pictureFormat.resetSize' },
      ],
      state: { kind: 'value' },
    },
    {
      id: 'styles',
      shape: 'dropdown',
      labelKey: 'pictureFormat.styles',
      toggle: 'choice',
      paths: icon('style'),
      options: PICTURE_STYLE_IDS.map((s) => ({
        value: s,
        labelKey: `pictureFormat.style.${s}` as TranslationKey,
      })),
      state: { kind: 'value' },
    },
    {
      id: 'altText',
      labelKey: 'pictureFormat.altText',
      paths: icon('accessibility_new'),
      state: { kind: 'dialog', dialog: 'imageAltText' },
    },
    {
      id: 'crop',
      labelKey: 'pictureFormat.crop',
      toggle: 'pressed',
      paths: icon('crop'),
      state: { kind: 'command' },
    },
    {
      id: 'cropToShape',
      defaultToolbar: false,
      shape: 'gallery',
      labelKey: 'pictureFormat.cropToShape',
      paths: icon('shapes'),
      state: { kind: 'value' },
    },
    {
      id: 'cropFit',
      defaultToolbar: false,
      shape: 'dropdown',
      labelKey: 'pictureFormat.cropFit',
      paths: icon('frame_person'),
      options: CROP_FITS.map((v) =>
        v === 'fill' || v === 'fit'
          ? { value: v, labelKey: `pictureFormat.${v}` as const }
          : { value: v, labelKey: 'pictureFormat.aspect', labelParams: { ratio: v } }
      ),
      state: { kind: 'value' },
    },
    {
      id: 'properties',
      labelKey: 'pictureFormat.properties',
      paths: icon('tune'),
      state: { kind: 'dialog', dialog: 'imageProperties' },
    },
  ] satisfies ChromeControl[],
};

/** The `image` slot ids. @public */
export type ImageSlotId =
  | 'image.corrections'
  | 'image.color'
  | 'image.transparency'
  | 'image.change'
  | 'image.reset'
  | 'image.styles'
  | 'image.altText'
  | 'image.crop'
  | 'image.cropToShape'
  | 'image.cropFit'
  | 'image.properties';

const isImageSlot = (slot: ChromeSlotId): slot is ImageSlotId => slot.startsWith('image.');

/**
 * The command an `image` value slot dispatches for a picked value, null when the value is
 * invalid, or undefined for a slot that is not an `image` value slot. @internal
 */
export function imageSlotCommand(
  slot: ChromeSlotId,
  value: unknown
): EditorCommand | null | undefined {
  if (!isImageSlot(slot)) return undefined;
  // A picked value is a string or a number; anything else (an object, null) is no choice.
  if (typeof value !== 'string' && typeof value !== 'number') return null;
  const text = String(value);
  switch (slot) {
    case 'image.corrections': {
      const [b, c] = text.split(',').map(Number);
      return Number.isFinite(b) && Number.isFinite(c)
        ? { type: 'setImageAdjustments', adjustments: { brightness: b / 100, contrast: c / 100 } }
        : null;
    }
    case 'image.color':
      return RECOLOR_PRESETS.includes(text as ImageRecolor)
        ? { type: 'setImageAdjustments', adjustments: { recolor: text as ImageRecolor } }
        : null;
    case 'image.transparency': {
      const t = Number(text);
      return Number.isFinite(t)
        ? { type: 'setImageAdjustments', adjustments: { transparency: t / 100 } }
        : null;
    }
    case 'image.reset':
      return text === 'picture' || text === 'size'
        ? { type: 'resetImage', ...(text === 'size' ? { size: true } : {}) }
        : null;
    case 'image.styles':
      return PICTURE_STYLE_IDS.includes(text as ImageStyleId)
        ? { type: 'setImageStyle', style: text as ImageStyleId }
        : null;
    case 'image.cropToShape':
      return text ? { type: 'cropToShape', preset: text } : null;
    case 'image.cropFit': {
      if (text === 'fill' || text === 'fit') return { type: 'cropImage', crop: { kind: text } };
      const [w, h] = text.split(':').map(Number);
      return w > 0 && h > 0
        ? { type: 'cropImage', crop: { kind: 'aspect', width: w, height: h } }
        : null;
    }
    case 'image.crop':
      return { type: 'toggleCrop' };
    case 'image.altText':
      // As DOCX Editor's `image.altText`: a picked value is the new description.
      return typeof value === 'string' ? { type: 'setAltText', description: value } : null;
    default:
      return null;
  }
}

/** A tiny image, only to ask `Editor.can` whether Change Picture can run. */
const PROBE_IMAGE: EditorCommand = {
  type: 'replaceImage',
  bytes: new Uint8Array([137, 80, 78, 71]),
  mime: 'image/png',
  width: 1,
  height: 1,
};

/** Probe values: well-formed picks, only to ask `Editor.can`. */
const PROBES: Partial<Record<ImageSlotId, unknown>> = {
  'image.corrections': '0,0',
  'image.color': 'none',
  'image.transparency': '0',
  'image.reset': 'picture',
  'image.styles': 'simpleFrameWhite',
  'image.cropToShape': 'ellipse',
  'image.cropFit': '1:1',
  'image.crop': 'toggle',
};

/**
 * The state of an `image` slot, or undefined for another slot. Every control needs one
 * selected picture, except Alt Text, which any one element takes. @internal
 */
export function imageSlotState(editor: Editor, id: ChromeSlotId): ToolbarCommandState | undefined {
  if (!isImageSlot(id)) return undefined;
  const snap = editor.snapshot();
  const pic = snap.image ?? null;
  const command =
    id === 'image.change'
      ? // An image with the frame's own aspect: a picture locked against resizing still takes
        // one, so the control stays on and the engine judges the real image when it applies.
        pic && pic.widthEmu > 0 && pic.heightEmu > 0
        ? { ...PROBE_IMAGE, width: pic.widthEmu, height: pic.heightEmu }
        : PROBE_IMAGE
      : id === 'image.altText'
        ? { type: 'setAltText' as const, description: '' }
        : id === 'image.properties'
          ? // A probe: `can` changes nothing. The current rotation is no change, so a picture
            // locked against rotation (`noRot`) still opens Format Picture.
            { type: 'setImageProperties' as const, rotationDegrees: pic?.rotationDegrees ?? 0 }
          : imageSlotCommand(id, PROBES[id]);
  const can = command ? editor.can(command) : { ok: false as const, reason: 'disabled.notWired' };
  const needsPicture = id !== 'image.altText' && can.ok && !pic;
  const reason = needsPicture ? 'disabled.noPicture' : can.ok ? null : can.reason;
  const value =
    id === 'image.altText'
      ? snap.geometry?.altText?.description
      : !pic
        ? undefined
        : id === 'image.corrections'
          ? correctionValue(Math.round(pic.brightness * 100), Math.round(pic.contrast * 100))
          : id === 'image.color'
            ? pic.recolor
            : id === 'image.transparency'
              ? String(Math.round(pic.transparency * 100))
              : id === 'image.styles'
                ? pic.style
                : id === 'image.cropToShape'
                  ? pic.shape
                  : undefined;
  return {
    id,
    enabled: reason === null,
    disabledReason: reason,
    active: id === 'image.crop' && !!pic?.cropping,
    ...(value !== undefined ? { value } : {}),
  };
}
