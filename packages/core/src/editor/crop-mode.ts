// Crop mode: the Crop button. While it is on, the surface shows the whole image and
// the crop handles of the selected picture (`surface-crop.ts`); a handle drag crops it with
// `cropImage`, and one drag is one undo step. Crop mode is editor state, not document state,
// like the format painter: a second press, Escape, or another selection ends it.

import type { EditorCommand, EditorSelection } from '../contracts/index.ts';
import { findItem } from '../automation/commands/placement.ts';
import { selectedIds, type CommandContext } from './commands.ts';
import { CommandRefusal } from '../automation/commands/op-batch.ts';
import { nodeIndex, xmlPartOf } from '../store/index.ts';
import { hasCropFillMode, pictureLock } from './picture-commands.ts';

/** The crop mode of one editor. */
export class CropMode {
  #target: { readonly part: string; readonly nodeId: string } | null = null;

  /** The state `Editor.can` answers depend on: a new object on each change. */
  get state(): object | null {
    return this.#target;
  }

  /**
   * The picture in crop mode on `part`, or null. Crop mode ends once the selection is no longer
   * that picture alone, so selecting the picture again does not bring it back.
   */
  nodeFor(part: string | undefined, selection: EditorSelection): string | null {
    const t = this.#target;
    if (!t) return null;
    const ids = selection.kind === 'shapes' ? selection.nodeIds : [];
    if (t.part === part && ids.length === 1 && ids[0] === t.nodeId) return t.nodeId;
    this.#target = null;
    return null;
  }

  /**
   * Run `toggleCrop`: `state` when it turned crop mode on or off (a probe changes nothing), null
   * for any other command. Throws CommandRefusal.
   */
  compile(ctx: CommandContext, command: EditorCommand): 'state' | null {
    if (command.type !== 'toggleCrop') return null;
    // As the picture commands: a field the command does not know is refused.
    if (Object.keys(command).some((k) => k !== 'type'))
      throw new CommandRefusal('errors.unsupported', 'toggleCrop');
    if (ctx.readOnly) throw new CommandRefusal('disabled.viewing');
    const part = ctx.session.info.slides[ctx.slide]?.partName;
    if (!part) throw new CommandRefusal('disabled.noEditor');
    const ids = selectedIds(ctx.selection);
    if (ids.length === 0) throw new CommandRefusal('disabled.noSelection');
    const items = ctx.session.layout(ctx.slide)?.items ?? [];
    if (ids.some((id) => findItem(items, id)?.kind !== 'picture'))
      throw new CommandRefusal('disabled.noPicture');
    if (ids.length !== 1) throw new CommandRefusal('disabled.onePicture');
    // A picture locked against cropping (`a:picLocks/@noCrop`) has no crop mode.
    const xml = xmlPartOf(ctx.session.store.package, part);
    const node = xml ? nodeIndex(xml).get(ids[0])?.node : undefined;
    if (node?.type === 'element' && pictureLock(node, 'noCrop'))
      throw new CommandRefusal('disabled.cropLocked');
    // Crop mode draws the image stretched over the frame: a tiled or inset fill has none.
    if (node?.type === 'element' && hasCropFillMode(node))
      throw new CommandRefusal('disabled.cropFillMode');
    if (ctx.probe) return 'state';
    // A second press on the picture in crop mode ends it.
    this.#target = this.nodeFor(part, ctx.selection) ? null : { part, nodeId: ids[0] };
    return 'state';
  }

  /**
   * After a command: Crop > Fill, Fit, and Aspect Ratio leave crop mode on, so the crop frame
   * shows for adjusting.
   */
  after(command: EditorCommand, part: string | undefined, selection: EditorSelection): void {
    if (command.type !== 'cropImage' || command.crop.kind === 'rect' || !part) return;
    const ids = selection.kind === 'shapes' ? selection.nodeIds : [];
    if (ids.length === 1) this.#target = { part, nodeId: ids[0] };
  }

  /** Turn crop mode off (Escape). True when it was on. */
  cancel(): boolean {
    if (!this.#target) return false;
    this.#target = null;
    return true;
  }
}
