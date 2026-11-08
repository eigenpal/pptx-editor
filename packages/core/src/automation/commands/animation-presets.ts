// The effects that `addAnimation` writes (ECMA-376 Part 1, §19.5): for each preset, its
// `presetID`, `presetClass`, and `presetSubtype`, and the behaviors that play it. The behaviors
// are the time-node elements of §19.5: `p:set` of `style.visibility` shows or hides the shape,
// `p:animEffect` runs a filter (`fade`, `wipe(up)`), `p:anim` moves or sizes it through the
// `ppt_x`, `ppt_y`, `ppt_w`, and `ppt_h` properties, `p:animScale` scales it, and `p:animRot`
// turns it. A presentation application reads the preset attributes to name the effect, and the
// behaviors to play it.

import type { AnimationClass, AnimationPreset } from '../../contracts/neutral.ts';

interface Preset {
  readonly effectClass: 'entrance' | 'emphasis' | 'exit';
  readonly presetId: number;
  readonly subtype: number;
  /** The duration of a new effect, in ms; null for an instant effect. */
  readonly duration: number | null;
  /** The behaviors, for the shape `spid` and a duration of `dur` ms. */
  readonly behaviors: (b: Behaviors, dur: number) => string;
}

/** The behavior writers of one effect: each takes the next `p:cTn/@id`. */
interface Behaviors {
  show(): string;
  hide(at: number): string;
  filter(transition: 'in' | 'out', filter: string, dur: number): string;
  anim(attr: string, dur: number, from: string, to: string): string;
  scale(dur: number, by: number, autoRev?: boolean): string;
  rotate(dur: number, by: number): string;
  pulse(dur: number): string;
}

/** The presets of `addAnimation`, in the order a menu offers them. @internal */
export const ANIMATION_PRESETS: Readonly<Record<AnimationPreset, Preset>> = {
  appear: {
    effectClass: 'entrance',
    presetId: 1,
    subtype: 0,
    duration: null,
    behaviors: (b) => b.show(),
  },
  fade: {
    effectClass: 'entrance',
    presetId: 10,
    subtype: 0,
    duration: 500,
    behaviors: (b, dur) => b.show() + b.filter('in', 'fade', dur),
  },
  flyIn: {
    effectClass: 'entrance',
    presetId: 2,
    subtype: 4,
    duration: 500,
    behaviors: (b, dur) =>
      b.show() +
      b.anim('ppt_x', dur, '#ppt_x', '#ppt_x') +
      b.anim('ppt_y', dur, '1+#ppt_h/2', '#ppt_y'),
  },
  wipe: {
    effectClass: 'entrance',
    presetId: 22,
    subtype: 4,
    duration: 500,
    behaviors: (b, dur) => b.show() + b.filter('in', 'wipe(up)', dur),
  },
  zoom: {
    effectClass: 'entrance',
    presetId: 53,
    subtype: 16,
    duration: 500,
    behaviors: (b, dur) =>
      b.show() +
      b.anim('ppt_w', dur, '0', '#ppt_w') +
      b.anim('ppt_h', dur, '0', '#ppt_h') +
      b.filter('in', 'fade', dur),
  },
  pulse: {
    effectClass: 'emphasis',
    presetId: 26,
    subtype: 0,
    duration: 500,
    behaviors: (b, dur) => b.pulse(dur) + b.scale(Math.max(1, Math.round(dur / 2)), 105000, true),
  },
  spin: {
    effectClass: 'emphasis',
    presetId: 8,
    subtype: 0,
    duration: 2000,
    behaviors: (b, dur) => b.rotate(dur, 21600000),
  },
  growShrink: {
    effectClass: 'emphasis',
    presetId: 6,
    subtype: 0,
    duration: 2000,
    behaviors: (b, dur) => b.scale(dur, 150000),
  },
  disappear: {
    effectClass: 'exit',
    presetId: 1,
    subtype: 0,
    duration: null,
    behaviors: (b) => b.hide(0),
  },
  fadeOut: {
    effectClass: 'exit',
    presetId: 10,
    subtype: 0,
    duration: 500,
    behaviors: (b, dur) => b.filter('out', 'fade', dur) + b.hide(dur - 1),
  },
  flyOut: {
    effectClass: 'exit',
    presetId: 2,
    subtype: 4,
    duration: 500,
    behaviors: (b, dur) =>
      b.anim('ppt_x', dur, '#ppt_x', '#ppt_x') +
      b.anim('ppt_y', dur, '#ppt_y', '1+#ppt_h/2') +
      b.hide(dur - 1),
  },
  wipeOut: {
    effectClass: 'exit',
    presetId: 22,
    subtype: 4,
    duration: 500,
    behaviors: (b, dur) => b.filter('out', 'wipe(down)', dur) + b.hide(dur - 1),
  },
  zoomOut: {
    effectClass: 'exit',
    presetId: 53,
    subtype: 16,
    duration: 500,
    behaviors: (b, dur) =>
      b.anim('ppt_w', dur, '#ppt_w', '0') +
      b.anim('ppt_h', dur, '#ppt_h', '0') +
      b.filter('out', 'fade', dur) +
      b.hide(dur - 1),
  },
};

/** Every preset id, for validation. @internal */
export const PRESET_IDS: ReadonlySet<string> = new Set(Object.keys(ANIMATION_PRESETS));

const PRESET_CLASS: Readonly<Record<Preset['effectClass'], string>> = {
  entrance: 'entr',
  emphasis: 'emph',
  exit: 'exit',
};

/** The preset of an effect with this class and `presetID`, or null. */
export function presetOf(
  effectClass: AnimationClass,
  presetId: number | null
): AnimationPreset | null {
  for (const [name, p] of Object.entries(ANIMATION_PRESETS))
    if (p.effectClass === effectClass && p.presetId === presetId) return name as AnimationPreset;
  return null;
}

/** What one new effect needs: its ids, its shape, and its timing. */
export interface EffectOptions {
  /** The next free `p:cTn/@id`. */
  readonly nextId: () => number;
  /** The `p:cNvPr/@id` of the shape; digits only. */
  readonly spid: string;
  readonly grpId: number;
  /** The `nodeType`: `clickEffect`, `withEffect`, or `afterEffect`. */
  readonly nodeType: string;
  readonly delayMs: number;
}

/** The XML of one new effect `p:par` of `preset`. */
export function effectXml(preset: AnimationPreset, o: EffectOptions): string {
  const p = ANIMATION_PRESETS[preset];
  const tgt = `<p:tgtEl><p:spTgt spid="${o.spid}"/></p:tgtEl>`;
  const visibility = (value: 'visible' | 'hidden', at: number): string =>
    `<p:set><p:cBhvr><p:cTn id="${o.nextId()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="${at}"/></p:stCondLst></p:cTn>` +
    `${tgt}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>` +
    `<p:to><p:strVal val="${value}"/></p:to></p:set>`;
  const b: Behaviors = {
    show: () => visibility('visible', 0),
    hide: (at) => visibility('hidden', Math.max(0, at)),
    filter: (transition, filter, dur) =>
      `<p:animEffect transition="${transition}" filter="${filter}"><p:cBhvr><p:cTn id="${o.nextId()}" dur="${dur}"/>${tgt}</p:cBhvr></p:animEffect>`,
    anim: (name, dur, from, to) =>
      `<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="${o.nextId()}" dur="${dur}" fill="hold"/>${tgt}` +
      `<p:attrNameLst><p:attrName>${name}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst>` +
      `<p:tav tm="0"><p:val><p:strVal val="${from}"/></p:val></p:tav>` +
      `<p:tav tm="100000"><p:val><p:strVal val="${to}"/></p:val></p:tav></p:tavLst></p:anim>`,
    scale: (dur, by, autoRev) =>
      `<p:animScale><p:cBhvr><p:cTn id="${o.nextId()}" dur="${dur}"${autoRev ? ' autoRev="1"' : ''} fill="hold"/>${tgt}</p:cBhvr>` +
      `<p:by x="${by}" y="${by}"/></p:animScale>`,
    rotate: (dur, by) =>
      `<p:animRot by="${by}"><p:cBhvr><p:cTn id="${o.nextId()}" dur="${dur}" fill="hold"/>${tgt}` +
      '<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>',
    pulse: (dur) =>
      `<p:animEffect transition="out" filter="fade"><p:cBhvr><p:cTn id="${o.nextId()}" dur="${dur}" tmFilter="0, 0; .2, .5; .8, .5; 1, 0"/>${tgt}</p:cBhvr>` +
      '<p:progress><p:fltVal val="0.25"/></p:progress></p:animEffect>',
  };
  const id = o.nextId();
  return (
    `<p:par><p:cTn id="${id}" presetID="${p.presetId}" presetClass="${PRESET_CLASS[p.effectClass]}" ` +
    `presetSubtype="${p.subtype}" fill="hold" grpId="${o.grpId}" nodeType="${o.nodeType}">` +
    `<p:stCondLst><p:cond delay="${o.delayMs}"/></p:stCondLst>` +
    `<p:childTnLst>${p.behaviors(b, p.duration ?? 1)}</p:childTnLst></p:cTn></p:par>`
  );
}
