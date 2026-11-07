// The chrome registry is public API (AGENTS.md, "React adapter"): `ChromeSlotId` values are
// stable names, adapters derive their toolbars from `CHROME_GROUPS`, and every label is an
// i18n key. These tests tie the three together so none can drift from the others.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import {
  CHROME_GROUPS,
  SHAPE_GALLERY,
  chromeControlForSlot,
  chromeSlotIds,
  chromeSlotIsToggle,
  commandForSlot,
  type ChromeSlotId,
} from '../index.ts';

/**
 * Every member of the `ChromeSlotId` union, once. `satisfies` fails to compile when the union
 * gains a member this record lacks, and the excess-property check fails when the record names
 * one the union does not have; the runtime test then ties the record to the registry.
 */
const SLOT_UNION = {
  'history.undo': true,
  'history.redo': true,
  'font.family': true,
  'font.size': true,
  'text.bold': true,
  'text.italic': true,
  'text.underline': true,
  'text.strike': true,
  'text.color': true,
  'text.highlight': true,
  'text.link': true,
  'format.clear': true,
  'text.characterSpacing': true,
  'text.changeCase': true,
  'text.dialog': true,
  'script.super': true,
  'script.sub': true,
  'paragraph.align': true,
  'list.bullet': true,
  'list.numbered': true,
  'list.bulletStyle': true,
  'list.numberStyle': true,
  'list.outdent': true,
  'list.indent': true,
  'list.lineSpacing': true,
  'paragraph.valign': true,
  'paragraph.dialog': true,
  'paragraph.textDirection': true,
  'paragraph.autofit': true,
  'paragraph.columns': true,
  'insert.textBox': true,
  'insert.line': true,
  'insert.shape': true,
  'insert.image': true,
  'shape.fill': true,
  'shape.lineColor': true,
  'shape.lineWidth': true,
  'shape.lineStart': true,
  'shape.lineEnd': true,
  'shape.lineDash': true,
  'shape.styles': true,
  'shape.effects': true,
  'shape.textEffects': true,
  'shape.properties': true,
  'format.painter': true,
  'arrange.order': true,
  'arrange.align': true,
  'arrange.group': true,
  'arrange.ungroup': true,
  'arrange.rotate': true,
  'arrange.selectionPane': true,
  'more.duplicate': true,
  'more.delete': true,
  'view.smartGuides': true,
  'view.presentationView': true,
  'find.find': true,
  'find.replace': true,
  'file.pageSetup': true,
  'edit.cut': true,
  'edit.copy': true,
  'edit.paste': true,
  'edit.selectAll': true,
  'slide.add': true,
  'slide.layout': true,
  'slide.reset': true,
  'slide.hide': true,
  'slide.background': true,
  'view.playFromStart': true,
  'view.playFromCurrent': true,
  'image.corrections': true,
  'image.color': true,
  'image.transparency': true,
  'image.change': true,
  'image.reset': true,
  'image.styles': true,
  'image.altText': true,
  'image.crop': true,
  'image.cropToShape': true,
  'image.cropFit': true,
  'image.properties': true,
  'section.add': true,
  'section.rename': true,
  'section.remove': true,
  'section.removeAll': true,
  'section.collapseAll': true,
  'section.expandAll': true,
  'table.insert': true,
  'table.style': true,
  'table.headerRow': true,
  'table.totalRow': true,
  'table.bandedRows': true,
  'table.firstColumn': true,
  'table.lastColumn': true,
  'table.bandedColumns': true,
  'table.cellFill': true,
  'table.borderTarget': true,
  'table.borderColor': true,
  'table.borderStyle': true,
  'table.borderWidth': true,
  'table.select': true,
  'table.delete': true,
  'table.insertRowAbove': true,
  'table.insertRowBelow': true,
  'table.insertColumnLeft': true,
  'table.insertColumnRight': true,
  'table.mergeCells': true,
  'table.splitCell': true,
  'table.distributeRows': true,
  'table.distributeColumns': true,
  'table.cellMargins': true,
} as const satisfies Record<ChromeSlotId, true>;

const en = JSON.parse(
  readFileSync(join(import.meta.dir, '..', '..', '..', '..', 'i18n', 'en.json'), 'utf8')
) as Record<string, unknown>;

/** Whether `key` names a string in en.json. */
function inCatalog(key: string): boolean {
  let node: unknown = en;
  for (const part of key.split('.')) {
    if (node === null || typeof node !== 'object' || !Object.hasOwn(node, part)) return false;
    node = (node as Record<string, unknown>)[part];
  }
  return typeof node === 'string';
}

describe('ChromeSlotId', () => {
  test('the union and chromeSlotIds() hold the same slots', () => {
    expect(chromeSlotIds().sort()).toEqual((Object.keys(SLOT_UNION) as ChromeSlotId[]).sort());
  });

  test('slot ids are unique', () => {
    const ids = chromeSlotIds();
    expect(new Set(ids).size).toBe(ids.length);
  });

  test('chromeControlForSlot finds every slot and nothing else', () => {
    for (const slot of chromeSlotIds()) expect(chromeControlForSlot(slot)).not.toBeNull();
    expect(chromeControlForSlot('text.missing' as ChromeSlotId)).toBeNull();
    expect(chromeControlForSlot('nodot' as ChromeSlotId)).toBeNull();
  });
});

describe('toggle flags', () => {
  test('the pressed toggles are exactly the slots whose command toggles a state', () => {
    for (const slot of chromeSlotIds()) {
      const command = commandForSlot(slot);
      // Hide Slide without `hidden` is the toggle of the current slide's state.
      const toggles =
        !!command &&
        (command.type.startsWith('toggle') ||
          (command.type === 'setSlidesHidden' && command.hidden === undefined));
      expect({ slot, toggle: chromeSlotIsToggle(slot) }).toEqual({ slot, toggle: toggles });
    }
  });

  test('a choice is a dropdown (a fixed option list, or the font picker)', () => {
    for (const group of CHROME_GROUPS)
      for (const control of group.controls)
        if (control.toggle === 'choice') expect(control.shape).toBe('dropdown');
  });
});

describe('i18n keys', () => {
  test('every group, control, and option label is in en.json', () => {
    const missing: string[] = [];
    for (const group of CHROME_GROUPS) {
      if (!inCatalog(group.labelKey)) missing.push(group.labelKey);
      for (const control of group.controls) {
        if (!inCatalog(control.labelKey)) missing.push(control.labelKey);
        for (const option of control.options ?? [])
          if (!inCatalog(option.labelKey)) missing.push(option.labelKey);
      }
    }
    expect(missing).toEqual([]);
  });

  test('every gallery category and preset has a label', () => {
    const missing: string[] = [];
    for (const category of SHAPE_GALLERY) {
      if (!inCatalog(category.labelKey)) missing.push(category.labelKey);
      for (const preset of category.presets)
        if (!inCatalog(`shapes.${preset}`)) missing.push(`shapes.${preset}`);
    }
    expect(missing).toEqual([]);
  });
});
