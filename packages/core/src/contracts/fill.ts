// The fill types of the public contract: gradients, patterns, pictures, and the slide
// background. Split from `neutral.ts`, which re-exports them, for the line cap.

import type { ColorChoice } from './neutral.ts';

/** A gradient fill as the chrome edits it. @public */
export interface GradientFill {
  /** `linear` (`a:lin`) or `radial` from the center (`a:path path="circle"`). */
  readonly kind: 'linear' | 'radial';
  /** Degrees clockwise from left-to-right (`a:lin/@ang`); 90 runs top to bottom. */
  readonly angle: number;
  /** Two to ten stops; `pos` from 0 to 1, `color` as `#RRGGBB`. */
  readonly stops: readonly { readonly pos: number; readonly color: string }[];
  /** The gradient turns with the shape (`rotWithShape`). */
  readonly rotWithShape: boolean;
}

/** The 54 preset patterns (`ST_PresetPatternVal`, ECMA-376 Part 1, 20.1.10.51). @public */
export type PatternPreset =
  | 'pct5'
  | 'pct10'
  | 'pct20'
  | 'pct25'
  | 'pct30'
  | 'pct40'
  | 'pct50'
  | 'pct60'
  | 'pct70'
  | 'pct75'
  | 'pct80'
  | 'pct90'
  | 'horz'
  | 'vert'
  | 'ltHorz'
  | 'ltVert'
  | 'dkHorz'
  | 'dkVert'
  | 'narHorz'
  | 'narVert'
  | 'dashHorz'
  | 'dashVert'
  | 'cross'
  | 'dnDiag'
  | 'upDiag'
  | 'ltDnDiag'
  | 'ltUpDiag'
  | 'dkDnDiag'
  | 'dkUpDiag'
  | 'wdDnDiag'
  | 'wdUpDiag'
  | 'dashDnDiag'
  | 'dashUpDiag'
  | 'diagCross'
  | 'smCheck'
  | 'lgCheck'
  | 'smGrid'
  | 'lgGrid'
  | 'dotGrid'
  | 'smConfetti'
  | 'lgConfetti'
  | 'horzBrick'
  | 'diagBrick'
  | 'solidDmnd'
  | 'openDmnd'
  | 'dotDmnd'
  | 'plaid'
  | 'sphere'
  | 'weave'
  | 'divot'
  | 'shingle'
  | 'wave'
  | 'trellis'
  | 'zigZag';

/** A preset pattern fill (`a:pattFill`) as the chrome edits it. @public */
export interface PatternFill {
  readonly preset: PatternPreset;
  /** The foreground (the pattern's ink), as `#RRGGBB`. */
  readonly foreground: string;
  /** The background, as `#RRGGBB`. */
  readonly background: string;
}

/** The fill commands of shapes beyond a color and a gradient. @public */
export type FillCommand =
  /** Shape fill: a preset pattern. */
  | { readonly type: 'setShapePattern'; readonly pattern: PatternFill }
  /**
   * Shape fill: a PNG, JPEG, or GIF picture, stretched over the shape, or tiled at its own
   * size when `tile` is true.
   */
  | {
      readonly type: 'setShapePicture';
      readonly bytes: Uint8Array;
      readonly mime: string;
      readonly tile?: boolean;
    };

/** A slide background the chrome sets. @public */
export type SlideBackground =
  | { readonly kind: 'solid'; readonly color: ColorChoice }
  | { readonly kind: 'gradient'; readonly gradient: GradientFill }
  | { readonly kind: 'pattern'; readonly pattern: PatternFill }
  | {
      readonly kind: 'picture';
      readonly bytes: Uint8Array;
      readonly mime: string;
      /** Tile the picture at its own size instead of stretching it over the slide. */
      readonly tile?: boolean;
    };

/** The current slide's background, as the chrome shows it. @public */
export interface SlideBackgroundState {
  /** `none` is white; `other` is a fill the chrome does not edit. */
  readonly kind: 'solid' | 'gradient' | 'pattern' | 'picture' | 'none' | 'other';
  readonly color?: string;
  readonly gradient?: GradientFill;
  readonly pattern?: PatternFill;
  /** True when the background comes from the layout or master, not the slide. */
  readonly inherited: boolean;
}
