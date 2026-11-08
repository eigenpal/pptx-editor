// The animation and transition contract: the effects of a slide (ECMA-376 Part 1, §19.5
// `p:timing`), its transition (§19.3.1 `p:transition`), and the commands that edit them.

/**
 * The class of an effect (`p:cTn/@presetClass`): entrance (`entr`), emphasis (`emph`), exit
 * (`exit`), motion path (`path`), or another class, such as a media call. @public
 */
export type AnimationClass = 'entrance' | 'emphasis' | 'exit' | 'motionPath' | 'other';

/**
 * The effects that `addAnimation` writes. Entrance: `appear`, `fade`, `flyIn`, `wipe`, `zoom`.
 * Emphasis: `pulse`, `spin`, `growShrink`. Exit: `disappear`, `fadeOut`, `flyOut`, `wipeOut`,
 * `zoomOut`. @public
 */
export type AnimationPreset =
  | 'appear'
  | 'fade'
  | 'flyIn'
  | 'wipe'
  | 'zoom'
  | 'pulse'
  | 'spin'
  | 'growShrink'
  | 'disappear'
  | 'fadeOut'
  | 'flyOut'
  | 'wipeOut'
  | 'zoomOut';

/**
 * When an effect starts (`p:cTn/@nodeType`): on a click (`clickEffect`), with the effect before
 * it (`withEffect`), or after the effect before it (`afterEffect`). @public
 */
export type AnimationTrigger = 'onClick' | 'withPrevious' | 'afterPrevious';

/** One effect of a slide, in the order the slide plays it. @public */
export interface AnimationEffect {
  /** The effect's `p:cTn/@id`: the handle that the animation commands take. */
  readonly id: string;
  /** The `p:cNvPr/@id` of the shape it animates. */
  readonly shapeId: string;
  /** The node id of that shape on the slide, or null when the slide has no such shape. */
  readonly nodeId: string | null;
  /** The name of that shape (`p:cNvPr/@name`), or an empty string. */
  readonly shapeName: string;
  readonly effectClass: AnimationClass;
  /** The preset, when it is one that `addAnimation` writes; otherwise null. */
  readonly preset: AnimationPreset | null;
  /** The `presetID` of the effect, or null. */
  readonly presetId: number | null;
  readonly trigger: AnimationTrigger;
  /**
   * The main sequence (`mainSeq`), or a sequence that a click on another shape starts
   * (`interactive`). Only main sequence effects can be edited.
   */
  readonly sequence: 'main' | 'interactive';
  /** For an interactive effect, the `p:cNvPr/@id` of the shape whose click starts it. */
  readonly triggerShapeId?: string;
  /** The name of that shape, or an empty string. */
  readonly triggerShapeName?: string;
  /**
   * The click that starts it in the main sequence: 0 for effects that play when the slide
   * starts, then 1, 2, and so on. Null for an interactive effect.
   */
  readonly step: number | null;
  /** The delay before the effect starts, in milliseconds. */
  readonly delayMs: number;
  /** How long the effect plays, in milliseconds; null for an effect that is instant. */
  readonly durationMs: number | null;
  /** The paragraphs it animates (`p:pRg`, 0-based and inclusive), when it animates text. */
  readonly paragraphs?: { readonly start: number; readonly end: number };
  /** Whether the animation commands can change it. */
  readonly editable: boolean;
}

/** The kinds of transition that `setTransition` writes. @public */
export type TransitionKind = 'fade' | 'push' | 'wipe' | 'split' | 'cut';

/** The transition of a slide (`p:transition`). @public */
export interface SlideTransitionState {
  /** The kind, `none` with no effect, or `other` for a kind that `setTransition` does not write. */
  readonly kind: TransitionKind | 'none' | 'other';
  /** How long it plays, in milliseconds (`p14:dur`, else the speed `spd`). */
  readonly durationMs: number;
  /** Whether a click moves to the next slide (`advClick`). */
  readonly advanceOnClick: boolean;
  /** The time after which the slide moves on by itself (`advTm`), or null. */
  readonly advanceAfterMs: number | null;
}

/** The animations and the transition of the current slide (`EditorSnapshot.animations`). @public */
export interface SlideAnimations {
  /** The effects, main sequence first, in the order the slide plays them. */
  readonly effects: readonly AnimationEffect[];
  /**
   * Whether `addAnimation` and the effect commands can edit this slide. A slide whose timing is
   * in markup-compatibility content (`mc:AlternateContent`) or whose main sequence has a
   * structure the commands do not write is read-only.
   */
  readonly editable: boolean;
  readonly transition: SlideTransitionState;
}

/** The animation commands. @public */
export type AnimationCommand =
  /**
   * Add an effect to each selected shape, at the end of the main sequence. A shape in a group
   * animates as its whole group. `trigger` defaults to `onClick`.
   */
  | {
      readonly type: 'addAnimation';
      readonly preset: AnimationPreset;
      readonly trigger?: AnimationTrigger;
    }
  /** Remove the effect `effectId` (`AnimationEffect.id`) of the current slide. */
  | { readonly type: 'removeAnimation'; readonly effectId: string }
  /** Move the effect `effectId` to position `to` of the main sequence (0-based). */
  | { readonly type: 'moveAnimation'; readonly effectId: string; readonly to: number }
  /**
   * Change when the effect `effectId` starts and how long it plays. Absent fields stay. Delay is
   * 0 to 3,600,000 ms; duration is 10 to 59,000 ms and is refused for an instant effect.
   */
  | {
      readonly type: 'setAnimationTiming';
      readonly effectId: string;
      readonly trigger?: AnimationTrigger;
      readonly delayMs?: number;
      readonly durationMs?: number;
    }
  /**
   * Set the transition of the current slide, or of the slides `indices`. Absent fields keep
   * their values. `kind: 'none'` removes the effect. Duration is 10 to 59,000 ms;
   * `advanceAfterMs` is 0 to 3,600,000 ms, or null to clear it.
   */
  | {
      readonly type: 'setTransition';
      readonly kind?: TransitionKind | 'none';
      readonly durationMs?: number;
      readonly advanceOnClick?: boolean;
      readonly advanceAfterMs?: number | null;
      readonly indices?: readonly number[];
    }
  /**
   * Show the order number of each effect on its shape on the canvas. A view setting, not an
   * edit: no undo step, never saved.
   */
  | { readonly type: 'showAnimationOrder'; readonly show: boolean };
