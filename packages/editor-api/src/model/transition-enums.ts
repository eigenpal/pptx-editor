/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The enums of `SlideTransition`, an SDK extension, and their ECMA-376 values.

import { enumMap } from './enums.ts';

/**
 * The effect of a slide transition: the effect element of `p:transition` (`CT_SlideTransition`,
 * ECMA-376 Part 1, §19.3.1.50), in PascalCase. A write takes `None`, `Fade`, `Push`, `Wipe`,
 * `Split`, and `Cut`. The other values are read states only: the other standard effect
 * elements, and `Other` for an effect from an extension namespace. An SDK extension: the pinned
 * Office.js reference has no such type.
 *
 * @public
 */
export enum SlideTransitionEffect {
  /** No effect element. */
  none = 'None',
  /** `p:fade`: the previous slide fades into the new slide. */
  fade = 'Fade',
  /** `p:push`: the new slide pushes the previous slide off the screen. */
  push = 'Push',
  /** `p:wipe`: the new slide wipes over the previous slide from one edge. */
  wipe = 'Wipe',
  /** `p:split`: the new slide shows from a split of the previous slide. */
  split = 'Split',
  /** `p:cut`: the new slide replaces the previous slide at once. */
  cut = 'Cut',
  /** `p:blinds`. Read only. */
  blinds = 'Blinds',
  /** `p:checker`. Read only. */
  checker = 'Checker',
  /** `p:circle`. Read only. */
  circle = 'Circle',
  /** `p:comb`. Read only. */
  comb = 'Comb',
  /** `p:cover`. Read only. */
  cover = 'Cover',
  /** `p:diamond`. Read only. */
  diamond = 'Diamond',
  /** `p:dissolve`. Read only. */
  dissolve = 'Dissolve',
  /** `p:newsflash`. Read only. */
  newsflash = 'Newsflash',
  /** `p:plus`. Read only. */
  plus = 'Plus',
  /** `p:pull`. Read only. */
  pull = 'Pull',
  /** `p:random`. Read only. */
  random = 'Random',
  /** `p:randomBar`. Read only. */
  randomBar = 'RandomBar',
  /** `p:strips`. Read only. */
  strips = 'Strips',
  /** `p:wedge`. Read only. */
  wedge = 'Wedge',
  /** `p:wheel`. Read only. */
  wheel = 'Wheel',
  /** `p:zoom`. Read only. */
  zoom = 'Zoom',
  /** An effect from an extension namespace. Read only. */
  other = 'Other',
}

/**
 * The effect elements of `CT_SlideTransition` (ECMA-376 Part 1, §19.3.1.50). A write names
 * §19.5.41 `fade`, §19.5.59 `push`, §19.5.96 `wipe`, §19.5.71 `split`, §19.5.34 `cut`, or `none`;
 * the other elements read only.
 */
export const TRANSITION_EFFECT = enumMap<`${SlideTransitionEffect}`>(
  [
    ['None', 'none'],
    ['Fade', 'fade'],
    ['Push', 'push'],
    ['Wipe', 'wipe'],
    ['Split', 'split'],
    ['Cut', 'cut'],
  ],
  {
    readOnly: [
      ['Blinds', 'blinds'],
      ['Checker', 'checker'],
      ['Circle', 'circle'],
      ['Comb', 'comb'],
      ['Cover', 'cover'],
      ['Diamond', 'diamond'],
      ['Dissolve', 'dissolve'],
      ['Newsflash', 'newsflash'],
      ['Plus', 'plus'],
      ['Pull', 'pull'],
      ['Random', 'random'],
      ['RandomBar', 'randomBar'],
      ['Strips', 'strips'],
      ['Wedge', 'wedge'],
      ['Wheel', 'wheel'],
      ['Zoom', 'zoom'],
      ['Other', 'other'],
    ],
    fallback: 'Other',
  }
);

/**
 * The direction of a slide transition effect: the side of `Push` and `Wipe`
 * (`ST_TransitionSideDirectionType`, ECMA-376 Part 1, §19.7.53), or the orientation
 * (`ST_Direction`, §19.7.2) and the direction (`ST_TransitionInOutDirectionType`, §19.7.52) of
 * `Split`. An SDK extension: the pinned Office.js reference has no such type.
 *
 * @public
 */
export enum SlideTransitionDirection {
  /** `dir="l"`, the schema default of `Push` and `Wipe`. */
  left = 'Left',
  /** `dir="u"`. */
  up = 'Up',
  /** `dir="r"`. */
  right = 'Right',
  /** `dir="d"`. */
  down = 'Down',
  /** `orient="horz"` with `dir="in"`. */
  horizontalIn = 'HorizontalIn',
  /** `orient="horz"` with `dir="out"`, the schema default of `Split`. */
  horizontalOut = 'HorizontalOut',
  /** `orient="vert"` with `dir="in"`. */
  verticalIn = 'VerticalIn',
  /** `orient="vert"` with `dir="out"`. */
  verticalOut = 'VerticalOut',
}

/** The directions of `SlideTransitionDirection` and their engine values. */
export const TRANSITION_DIRECTION = enumMap<`${SlideTransitionDirection}`>([
  ['Left', 'l'],
  ['Up', 'u'],
  ['Right', 'r'],
  ['Down', 'd'],
  ['HorizontalIn', 'horzIn'],
  ['HorizontalOut', 'horzOut'],
  ['VerticalIn', 'vertIn'],
  ['VerticalOut', 'vertOut'],
]);
