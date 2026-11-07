// The text format commands of the Font and Paragraph groups and their dialogs: Format >
// Paragraph, superscript and subscript, Change Case, character spacing, Format > Font, Bullets
// and Numbering, Text Direction, and Columns. Split from `commands.ts`,
// so the command table stays under the line cap.

import { child } from '../store/index.ts';
import type { EditorCommand, TextFormatCommand } from '../contracts/index.ts';
import { CommandRefusal, type OpBatch } from '../automation/commands/op-batch.ts';
import type { CommandContext, CommandEffect } from './commands.ts';
import {
  forRuns,
  listTarget,
  probeLists,
  refitIds,
  refitTargets,
  selectedIds,
  slidePart,
  textTargets,
} from './text-targets.ts';
import {
  applyParagraphFormat,
  checkParagraphFormat,
  onlyRemoves,
} from '../automation/commands/paragraph-format.ts';
import { setColumns, setTextDirection } from './text-body-commands.ts';
import { setBulletFormats } from './bullet-format.ts';
import { compilePasteText } from './clipboard-text.ts';
import {
  applyFontFormat,
  baselineValue,
  changeCase,
  checkCharacterSpacing,
  checkFontFormat,
  spacingValue,
  SUBSCRIPT_BASELINE,
  SUPERSCRIPT_BASELINE,
} from '../automation/commands/font-format.ts';
import {
  ensureParagraphProps,
  paragraphsForFormatting,
  setRunAttr,
} from '../automation/commands/text-model.ts';

const TYPES: ReadonlySet<string> = new Set([
  'setParagraphFormat',
  'toggleSuperscript',
  'toggleSubscript',
  'changeCase',
  'setCharacterSpacing',
  'setBulletFormat',
  'setTextDirection',
  'setColumns',
  'setFontFormat',
  'pasteText',
]);

export function isTextFormatCommand(command: EditorCommand): command is TextFormatCommand {
  return TYPES.has(command.type);
}

/** Compile a text format command into `batch`. Throws CommandRefusal. */
export function compileTextFormatCommand(
  ctx: CommandContext,
  batch: OpBatch,
  command: TextFormatCommand
): CommandEffect {
  switch (command.type) {
    case 'pasteText':
      return compilePasteText(ctx, batch, command);
    case 'setParagraphFormat': {
      checkParagraphFormat(command);
      const part = slidePart(ctx);
      // A command that only removes values leaves a paragraph with no `a:pPr` as it is.
      const removing = onlyRemoves(command);
      const targets = textTargets(batch, part, ctx.selection);
      for (const t of targets)
        for (const pId of paragraphsForFormatting(batch, part, t.txBodyId, t.range)) {
          const own = child(batch.element(part, pId), 'a:pPr');
          if (removing && !own) continue;
          const pPr = own?.id ?? ensureParagraphProps(batch, part, pId);
          applyParagraphFormat(batch, part, pPr, command);
        }
      refitTargets(batch, ctx, part, targets);
      return { label: 'paragraphFormat' };
    }
    case 'toggleSuperscript':
    case 'toggleSubscript': {
      const part = slidePart(ctx);
      const up = command.type === 'toggleSuperscript';
      const on = up ? !ctx.active.superscript : !ctx.active.subscript;
      const value = baselineValue(on ? (up ? SUPERSCRIPT_BASELINE : SUBSCRIPT_BASELINE) : 0);
      forRuns(batch, part, ctx, (run) => setRunAttr(batch, part, run, 'baseline', value));
      return { label: command.type };
    }
    case 'changeCase': {
      const part = slidePart(ctx);
      const targets = textTargets(batch, part, ctx.selection);
      for (const t of targets) changeCase(batch, part, t.txBodyId, t.range, command.case);
      refitTargets(batch, ctx, part, targets);
      return { label: 'changeCase' };
    }
    case 'setCharacterSpacing': {
      const part = slidePart(ctx);
      const spc = spacingValue(checkCharacterSpacing(command.points));
      forRuns(batch, part, ctx, (run) => setRunAttr(batch, part, run, 'spc', spc));
      return { label: 'characterSpacing' };
    }
    case 'setBulletFormat': {
      const part = slidePart(ctx);
      const targets = textTargets(batch, part, ctx.selection);
      // A format with a kind applies to every paragraph, so a check reads a few. One without
      // reads which paragraphs show a list: all of them.
      const lists = command.format?.kind
        ? probeLists(batch, part, ctx, targets)
        : targets.map((t) => listTarget(batch, part, t));
      const format = command.format;
      const { listed, changed } = setBulletFormats(
        batch,
        part,
        ctx.slide,
        lists,
        format,
        !ctx.probe
      );
      // Refit only a shape whose bullets changed, as `setColumns` does.
      refitTargets(batch, ctx, part, changed);
      // A size or a color alone changes only paragraphs that show a list: with none, the
      // command is refused (also for `can`), not an empty undo step.
      if (!listed) throw new CommandRefusal('disabled.noList');
      return { label: 'bulletFormat' };
    }
    case 'setTextDirection':
    case 'setColumns': {
      const part = slidePart(ctx);
      const ids = selectedIds(ctx.selection);
      const changed =
        command.type === 'setTextDirection'
          ? setTextDirection(batch, part, ids, command.direction)
          : setColumns(batch, part, ids, command.count, command.gapEmu);
      refitIds(batch, ctx, part, changed);
      return { label: command.type };
    }
    case 'setFontFormat': {
      checkFontFormat(command);
      const part = slidePart(ctx);
      forRuns(batch, part, ctx, (run) => applyFontFormat(batch, part, run, command));
      return { label: 'fontFormat' };
    }
  }
}
