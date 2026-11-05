// Field rules of the picture dialogs: Alt Text (the Alt Text pane) and Format Picture
// (size, rotation, and crop position). What each dialog opens with, what blocks OK, and the
// command OK applies. Framework-free, so the React and Vue dialogs agree. Every export is
// adapter glue, not a stable API: `@internal`.

import type {
  AltText,
  EditorCommand,
  ImageCropPercent,
  SelectedImageState,
} from '../contracts/index.ts';
import type { DialogErrors } from './dialog-draft.ts';
import { MAX_ALT_TEXT } from './picture-commands.ts';

export { MAX_ALT_TEXT };

/** The draft of the Alt Text dialog. @internal */
export interface AltTextDraftFields {
  readonly description: string;
  readonly title: string;
  readonly decorative: boolean;
}

/** The draft the dialog opens with: the selected element's alt text. @internal */
export function altTextDraftFrom(alt: AltText | undefined): AltTextDraftFields {
  return {
    description: alt?.description ?? '',
    title: alt?.title ?? '',
    decorative: alt?.decorative ?? false,
  };
}

/** Errors that block OK: text longer than the editor writes. A decorative element keeps no description, so its length does not count. @internal */
export function validateAltTextDraft(fields: AltTextDraftFields): DialogErrors<AltTextDraftFields> {
  return {
    ...(!fields.decorative && fields.description.length > MAX_ALT_TEXT
      ? { description: 'errors.altTextLength' }
      : {}),
    ...(fields.title.length > MAX_ALT_TEXT ? { title: 'errors.altTextLength' } : {}),
  };
}

/**
 * The command OK applies: only the fields that changed, so an edit by another participant to
 * another field, or a title the user did not touch, stays as it is. Null when nothing changed.
 * @internal
 */
export function altTextDraftCommand(
  fields: AltTextDraftFields,
  initial: AltTextDraftFields
): EditorCommand | null {
  const decorative = fields.decorative !== initial.decorative;
  const description = !fields.decorative && fields.description !== initial.description;
  const title = fields.title !== initial.title;
  if (!decorative && !description && !title) return null;
  return {
    type: 'setAltText',
    ...(description || (decorative && !fields.decorative)
      ? { description: fields.description }
      : {}),
    ...(title ? { title: fields.title } : {}),
    ...(decorative ? { decorative: fields.decorative } : {}),
  };
}

const EMU_PER_INCH = 914_400;
/** The largest picture side, in inches. */
const MAX_SIDE_INCHES = 56;

/** The draft of the Format Picture dialog, in the units it shows: inches, degrees, percent. @internal */
export interface PicturePropertiesDraftFields {
  readonly width: number;
  readonly height: number;
  /** Lock aspect ratio: a width change scales the height, and the other way round. */
  readonly lockAspect: boolean;
  readonly rotation: number;
  readonly cropLeft: number;
  readonly cropTop: number;
  readonly cropRight: number;
  readonly cropBottom: number;
}

const CROP_FIELDS = ['cropLeft', 'cropTop', 'cropRight', 'cropBottom'] as const;
type CropField = (typeof CROP_FIELDS)[number];

const round = (v: number, digits: number): number => Math.round(v * 10 ** digits) / 10 ** digits;

/**
 * The draft the Format Picture dialog opens with, from the snapshot: the size on the slide, as
 * the Size fields show it. @internal
 */
export function picturePropertiesDraftFrom(
  picture:
    | (Pick<SelectedImageState, 'widthEmu' | 'heightEmu' | 'rotationDegrees' | 'crop'> &
        Partial<Pick<SelectedImageState, 'lockAspect'>>)
    | null
    | undefined
): PicturePropertiesDraftFields {
  const crop = picture?.crop ?? { left: 0, top: 0, right: 0, bottom: 0 };
  return {
    width: round((picture?.widthEmu ?? 0) / EMU_PER_INCH, 2),
    height: round((picture?.heightEmu ?? 0) / EMU_PER_INCH, 2),
    // As DOCX Editor: the picture's own lock (`noChangeAspect`), on by default.
    lockAspect: picture?.lockAspect ?? true,
    rotation: round(picture?.rotationDegrees ?? 0, 1),
    cropLeft: round(crop.left, 1),
    cropTop: round(crop.top, 1),
    cropRight: round(crop.right, 1),
    cropBottom: round(crop.bottom, 1),
  };
}

/**
 * The draft after the user types `value` into `field`: with Lock aspect ratio on, the other
 * side follows at the aspect ratio the dialog opened with (`initial`). @internal
 */
export function picturePropertiesWithValue(
  fields: PicturePropertiesDraftFields,
  field: 'width' | 'height',
  value: number,
  initial: PicturePropertiesDraftFields
): PicturePropertiesDraftFields {
  const other = field === 'width' ? 'height' : 'width';
  const scale = initial[field] > 0 && Number.isFinite(value) ? value / initial[field] : NaN;
  return fields.lockAspect && Number.isFinite(scale)
    ? { ...fields, [field]: value, [other]: round(initial[other] * scale, 2) }
    : { ...fields, [field]: value };
}

/**
 * Errors that block OK: a size, angle, or crop the picture cannot take. With `initial`, a side
 * the user did not change is not checked, so a picture outside the size range still
 * takes a new rotation or crop. @internal
 */
export function validatePicturePropertiesDraft(
  f: PicturePropertiesDraftFields,
  initial?: PicturePropertiesDraftFields
): DialogErrors<PicturePropertiesDraftFields> {
  const inRange = (v: number): boolean => Number.isFinite(v) && v >= 0.01 && v <= MAX_SIDE_INCHES;
  const side = (v: number, was: number | undefined): boolean => v === was || inRange(v);
  // A crop edge outside the range the dialog takes; an unchanged edge keeps what the picture has
  // (a Fit crop can pad far more than the dialog offers).
  const pad = (field: CropField): boolean =>
    f[field] !== initial?.[field] &&
    !(Number.isFinite(f[field]) && f[field] > -1000 && f[field] < 100);
  const errors: Partial<Record<keyof PicturePropertiesDraftFields, string>> = {};
  if (!side(f.width, initial?.width)) errors.width = 'imageProperties.sizeRange';
  if (!side(f.height, initial?.height)) errors.height = 'imageProperties.sizeRange';
  if (!Number.isFinite(f.rotation)) errors.rotation = 'imageProperties.rotationRange';
  for (const field of CROP_FIELDS) if (pad(field)) errors[field] = 'imageProperties.cropPadRange';
  // A changed pair that leaves nothing visible marks its second side; an unchanged pair stays.
  const changed = (a: CropField, b: CropField): boolean =>
    f[a] !== initial?.[a] || f[b] !== initial?.[b];
  if (
    !errors.cropRight &&
    !errors.cropLeft &&
    changed('cropLeft', 'cropRight') &&
    !(f.cropLeft + f.cropRight < 99.9)
  )
    errors.cropRight = 'imageProperties.cropRange';
  if (
    !errors.cropBottom &&
    !errors.cropTop &&
    changed('cropTop', 'cropBottom') &&
    !(f.cropTop + f.cropBottom < 99.9)
  )
    errors.cropBottom = 'imageProperties.cropRange';
  return errors;
}

/**
 * The command OK applies: only the values that changed, or null when none did. An edge whose
 * field did not change keeps its exact crop (`exact`, the picture's crop when the dialog opened),
 * not the rounded value the field shows. @internal
 */
export function picturePropertiesDraftCommand(
  f: PicturePropertiesDraftFields,
  initial: PicturePropertiesDraftFields,
  exact?: ImageCropPercent
): EditorCommand | null {
  const edge = (now: number, was: number, original: number | undefined): number =>
    now === was && original !== undefined ? original : now;
  const cropChanged =
    f.cropLeft !== initial.cropLeft ||
    f.cropTop !== initial.cropTop ||
    f.cropRight !== initial.cropRight ||
    f.cropBottom !== initial.cropBottom;
  const properties = {
    ...(f.width !== initial.width ? { widthEmu: Math.round(f.width * EMU_PER_INCH) } : {}),
    ...(f.height !== initial.height ? { heightEmu: Math.round(f.height * EMU_PER_INCH) } : {}),
    ...(f.rotation !== initial.rotation ? { rotationDegrees: f.rotation } : {}),
    ...(cropChanged
      ? {
          crop: {
            left: edge(f.cropLeft, initial.cropLeft, exact?.left),
            top: edge(f.cropTop, initial.cropTop, exact?.top),
            right: edge(f.cropRight, initial.cropRight, exact?.right),
            bottom: edge(f.cropBottom, initial.cropBottom, exact?.bottom),
          },
        }
      : {}),
  };
  return Object.keys(properties).length ? { type: 'setImageProperties', ...properties } : null;
}
