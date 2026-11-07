// The animation and transition commands. `commands.ts` sends them here; each compiles to one
// transaction, so one undo step. The edits are in `automation/commands/animation-commands.ts`
// and `transition-commands.ts`.

import type { EditorCommand, SlideAnimations } from '../contracts/index.ts';
import { presentationInfo, xmlPartOf, type Element, type PptxPackage } from '../store/index.ts';
import {
  addAnimation,
  moveAnimation,
  removeAnimation,
  setAnimationTiming,
} from '../automation/commands/animation-commands.ts';
import { readAnimations } from '../automation/commands/animation-model.ts';
import { readTransition, setTransition } from '../automation/commands/transition-commands.ts';
import { CommandRefusal, type OpBatch } from '../automation/commands/op-batch.ts';
import { slideIndices } from '../automation/commands/slide-multi.ts';
import type { CommandContext, CommandEffect } from './commands.ts';
import { slidePart } from './text-targets.ts';

/** The commands this table compiles. */
export type AnimationEditCommand = Extract<
  EditorCommand,
  {
    readonly type:
      'addAnimation' | 'removeAnimation' | 'moveAnimation' | 'setAnimationTiming' | 'setTransition';
  }
>;

const TYPES: ReadonlySet<string> = new Set([
  'addAnimation',
  'removeAnimation',
  'moveAnimation',
  'setAnimationTiming',
  'setTransition',
]);

export function isAnimationCommand(command: EditorCommand): command is AnimationEditCommand {
  return TYPES.has(command.type);
}

/** The shapes an effect goes on: the selected shapes, or the shape of the text or cells. */
function selectedNodes(ctx: CommandContext): readonly string[] {
  const sel = ctx.selection;
  if (sel.kind === 'shapes') return sel.nodeIds;
  if (sel.kind === 'none') return [];
  return [sel.nodeId];
}

export function compileAnimationCommand(
  ctx: CommandContext,
  batch: OpBatch,
  command: AnimationEditCommand
): CommandEffect {
  if (command.type === 'setTransition') {
    const slides = presentationInfo(batch.pkg).slides;
    const targets = command.indices ? slideIndices(batch, command.indices) : [ctx.slide];
    let changed = 0;
    for (const i of targets) {
      const part = slides[i]?.partName;
      if (!part) throw new CommandRefusal('errors.malformed');
      try {
        setTransition(batch, part, command);
        changed += 1;
      } catch (e) {
        // A slide that has the transition already is skipped; the others change.
        if (!(e instanceof CommandRefusal && e.reason === 'disabled.noAnimationChange')) throw e;
      }
    }
    if (changed === 0) throw new CommandRefusal('disabled.noAnimationChange');
    return { label: 'transition', selection: ctx.selection };
  }
  const part = slidePart(ctx);
  switch (command.type) {
    case 'addAnimation': {
      const nodes = selectedNodes(ctx);
      if (nodes.length === 0) throw new CommandRefusal('disabled.noSelection');
      addAnimation(batch, part, nodes, command.preset, command.trigger);
      break;
    }
    case 'removeAnimation':
      removeAnimation(batch, part, String(command.effectId));
      break;
    case 'moveAnimation':
      moveAnimation(batch, part, String(command.effectId), command.to);
      break;
    case 'setAnimationTiming':
      setAnimationTiming(batch, part, String(command.effectId), command);
      break;
  }
  return { label: 'animation', selection: ctx.selection };
}

const EMPTY: SlideAnimations = Object.freeze({
  effects: Object.freeze([]),
  editable: false,
  transition: Object.freeze({
    kind: 'none',
    durationMs: 750,
    advanceOnClick: true,
    advanceAfterMs: null,
  }),
});

/**
 * The snapshot source of `EditorSnapshot.animations`: the effects and the transition of a slide
 * part, reference-stable while the part's root does not change.
 */
export function createAnimationsSource(): (
  pkg: PptxPackage | undefined,
  part: string | undefined
) => SlideAnimations {
  let last: { root: Element; out: SlideAnimations } | undefined;
  return (pkg, part) => {
    const root = pkg && part ? xmlPartOf(pkg, part)?.root : undefined;
    if (!root) return EMPTY;
    if (last?.root === root) return last.out;
    const { effects, editable } = readAnimations(root);
    const out: SlideAnimations = { effects, editable, transition: readTransition(root) };
    last = { root, out };
    return out;
  };
}
