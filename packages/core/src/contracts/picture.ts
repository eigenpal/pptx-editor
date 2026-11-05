// The picture types of the public contract: crop, adjustments, styles, the Format Picture
// fields, the selected picture's state, and alt text. Split from `index.ts`, which re-exports
// them, for the line cap.

/**
 * The crop of each edge of a picture, in percent of the image (`a:srcRect`), after DOCX
 * Editor's `ImageCropPercent`: a positive value removes that part of the image. Unlike DOCX
 * Editor's 0 to 100, a value may be negative: it adds space on that side as
 * padding (Fit writes it). @public
 */
export interface ImageCropPercent {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

/**
 * A picture crop. `rect` crops each edge in percent: the image stays in place and the frame
 * changes, as a crop handle does; with `keepFrame`, the frame stays and the image moves in it.
 * `fill` and `fit` keep the frame and crop the image to cover it or to show all of it; `aspect`
 * crops the image to a width-to-height ratio. @public
 */
export type ImageCrop =
  | ({ readonly kind: 'rect'; readonly keepFrame?: boolean } & ImageCropPercent)
  | { readonly kind: 'fill' }
  | { readonly kind: 'fit' }
  | { readonly kind: 'aspect'; readonly width: number; readonly height: number };

/**
 * The Recolor presets of the Color gallery. `accentNDark` and `accentNLight` are the
 * dark and light variations of theme accent N (`a:duotone`). @public
 */
export type ImageRecolor =
  | 'none'
  | 'grayscale'
  | 'sepia'
  | 'blackWhite25'
  | 'blackWhite50'
  | 'blackWhite75'
  | `accent${1 | 2 | 3 | 4 | 5 | 6}${'Dark' | 'Light'}`;

/**
 * The Format Picture values, as DOCX Editor's `setImageProperties` names them. The size is the
 * size on the slide, as the Size fields show it (`SelectedImageState.widthEmu`; inside a
 * scaled group it differs from the file's `a:ext`), and keeps the top-left corner of the unrotated box;
 * the crop keeps the image fixed on the slide. @public
 */
export interface ImageProperties {
  /** The width on the slide, in EMU. */
  readonly widthEmu?: number;
  /** The height on the slide, in EMU. */
  readonly heightEmu?: number;
  /** The rotation in degrees, clockwise. */
  readonly rotationDegrees?: number;
  /** The crop of each edge in percent of the image. */
  readonly crop?: ImageCropPercent;
}

/** Picture corrections, color, and transparency, as the Picture Format tab sets them. @public */
export interface ImageAdjustments {
  /** Brightness from −1 to 1 (`a:lum/@bright`). */
  readonly brightness?: number;
  /** Contrast from −1 to 1 (`a:lum/@contrast`). */
  readonly contrast?: number;
  /** Transparency from 0 to 1 (`a:alphaModFix`, 1 − amount). */
  readonly transparency?: number;
  /** A Recolor preset of the Color gallery; `none` removes the color effect. */
  readonly recolor?: ImageRecolor;
}

/** The picture styles the gallery offers, with the gallery's names in camel case. @public */
export type ImageStyleId =
  | 'simpleFrameWhite'
  | 'simpleFrameBlack'
  | 'softEdgeRectangle'
  | 'softEdgeOval'
  | 'reflectedRoundedRectangle';

/**
 * The selected pictures, as the picture controls show them: the values are the first
 * picture's. @public
 */
export interface SelectedImageState {
  /** The first selected picture. */
  readonly nodeId: string;
  /** How many pictures are selected. Crop, Change Picture, Format Picture, and Alt Text take one. */
  readonly count: number;
  /** The width on the slide, in EMU, as `setImageProperties` takes it. */
  readonly widthEmu: number;
  /** The height on the slide, in EMU, as `setImageProperties` takes it. */
  readonly heightEmu: number;
  /** The rotation in degrees, clockwise. */
  readonly rotationDegrees: number;
  /** The crop of each edge in percent of the image (`a:srcRect`). */
  readonly crop: ImageCropPercent;
  /**
   * The preset the picture is cropped to (`rect` for none), or `custom` (`a:custGeom`); the
   * value `cropToShape` takes as `preset`.
   */
  readonly shape: string;
  /** Crop mode is on: the crop handles show. */
  readonly cropping: boolean;
  /** Brightness from −1 to 1 (`a:lum/@bright`); 0 is none. */
  readonly brightness: number;
  /** Contrast from −1 to 1 (`a:lum/@contrast`); 0 is none. */
  readonly contrast: number;
  /** Transparency from 0 (opaque) to 1 (`a:alphaModFix`, 1 − amount). */
  readonly transparency: number;
  /** The Recolor preset the picture matches, or `other` for another color effect. */
  readonly recolor: ImageRecolor | 'other';
  /** The picture style the picture matches, if any. */
  readonly style?: ImageStyleId;
  /**
   * The value of the Lock Aspect Ratio check box (`a:picLocks/@noChangeAspect`): the
   * Format Picture default. Unlike DOCX Editor's `locks.changeAspect`, the engine does not
   * enforce it, and the dialog does not save the box.
   */
  readonly lockAspect: boolean;
  /** Whether the picture can be cropped: false with `a:picLocks/@noCrop`, as DOCX Editor's `canCrop`. */
  readonly canCrop: boolean;
  /**
   * Whether `setImageProperties` and Reset Picture & Size can change the size: false with
   * `a:picLocks/@noResize`, as DOCX Editor's `canResize`.
   */
  readonly canResize: boolean;
  /** Whether `setImageProperties` can change the rotation: false with `a:picLocks/@noRot`. */
  readonly canRotate: boolean;
}

/** The alt text of an element (`p:cNvPr`). @public */
export interface AltText {
  readonly description: string;
  readonly title: string;
  /** Marked as decorative (the `adec:decorative` extension). */
  readonly decorative: boolean;
}

/**
 * The picture commands (`picture-commands.ts`): crop, crop to shape, crop mode, Change Picture,
 * Reset Picture, Corrections and Color, styles, Format Picture, and alt text. @public
 */
export type PictureCommand =
  /**
   * Crop the selected picture (`a:srcRect`): to a crop rectangle, where the image stays in place
   * on the slide and the frame changes; to fill or fit its frame; or to an aspect ratio.
   */
  | { readonly type: 'cropImage'; readonly crop: ImageCrop }
  /** Crop the selected pictures to a preset shape (`a:prstGeom`); `rect` crops back to a rectangle. */
  | { readonly type: 'cropToShape'; readonly preset: string }
  /**
   * The Crop button: turn crop mode on for the selected picture, or off. In crop mode the crop
   * handles show; Escape, a second press, or another selection ends it.
   */
  | { readonly type: 'toggleCrop' }
  /**
   * Change Picture: replace the selected picture's image with an image file. The new image fits
   * in the old frame; the crop resets, and the border, effects, and adjustments stay. `width` and
   * `height` are the stored pixel size; without them, the engine reads it from the image header.
   * A JPEG's EXIF orientation does not count: the stored pixels do.
   */
  | {
      readonly type: 'replaceImage';
      readonly bytes: Uint8Array;
      readonly mime: string;
      readonly width?: number;
      readonly height?: number;
    }
  /**
   * Reset Picture: remove the adjustments, the border, the effects, and the crop to a shape.
   * `size` also removes the crop and gives the picture its own size (Reset Picture & Size).
   */
  | { readonly type: 'resetImage'; readonly size?: boolean }
  /** Corrections, color, and transparency of the selected pictures. Absent fields stay. */
  | { readonly type: 'setImageAdjustments'; readonly adjustments: ImageAdjustments }
  /** A picture style of the Picture Styles gallery: border, effects, and shape. */
  | { readonly type: 'setImageStyle'; readonly style: ImageStyleId }
  /**
   * Format Picture: the size, rotation, and crop of the selected picture as one undo step, with
   * the fields on the command, after DOCX Editor's `setImageProperties`. It takes `widthEmu`,
   * `heightEmu`, and `crop` as DOCX Editor does, plus `rotationDegrees`; absent fields stay, and
   * any other field is refused. DOCX Editor's other fields live elsewhere: alt text in
   * `setAltText` (every object), links in `setLink`, the natural size in `resetImage` with
   * `size`, the border in the Shape line commands; it has no target check, as the dialog closes
   * when its picture changes. DOCX Editor's wrap and position fields have no slide equivalent.
   */
  | ({ readonly type: 'setImageProperties' } & ImageProperties)
  /**
   * Alt text of the selected element (`p:cNvPr/@descr` and `@title`). `decorative` marks it as
   * decorative, as the Alt Text pane does; a decorative element keeps no description,
   * and a description for one is refused unless `decorative` is false. An absent field keeps the
   * element's own; an empty description or title removes it.
   */
  | {
      readonly type: 'setAltText';
      readonly description?: string;
      readonly title?: string;
      readonly decorative?: boolean;
    };
