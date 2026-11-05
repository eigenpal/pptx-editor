// Custom document properties: the custom properties part (`docProps/custom.xml`, ECMA-376 Part 1,
// 22.3). Each `property` element has a name, a `pid`, and one variant child of the
// `docPropsVTypes` schema (22.4). The object model shows four types: String, Number, Boolean, and
// Date. A property of another variant type (a vector, a blob, a class id) is not listed, and
// `deleteAll` keeps it. An `add` with its key replaces it, as a key names one property.

import {
  attr,
  elements,
  isValidXmlText,
  textOf,
  type Element,
  type NodeSpec,
} from '../store/index.ts';
import type { AutomationCustomPropertyWrite } from './operations.ts';
import { editableRoot, ensurePart, packagePart, rootOf } from './plan-properties.ts';
import type { AutomationCustomPropertyRead, AutomationValue } from './protocol.ts';
import { AutomationFailure, badArgument, edit, type PlanContext } from './targets.ts';

const NS_CUSTOM = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties';
const NS_VT = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes';
const REL_CUSTOM =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties';
const CT_CUSTOM = 'application/vnd.openxmlformats-officedocument.custom-properties+xml';
/** The format id of user-defined properties (22.3.2.2 `property/@fmtid`). */
const FMTID = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}';
const CUSTOM_XML = `<Properties xmlns="${NS_CUSTOM}" xmlns:vt="${NS_VT}"/>`;

/** The longest key and string value a write accepts. */
export const MAX_CUSTOM_KEY = 255;
export const MAX_CUSTOM_VALUE = 32_767;

type CustomType = AutomationCustomPropertyRead['type'];

/** The variant elements each listed type reads from. Null prototype: file names are untrusted. */
const TYPE_OF: Readonly<Record<string, CustomType>> = Object.assign(Object.create(null), {
  lpwstr: 'String',
  lpstr: 'String',
  bstr: 'String',
  i1: 'Number',
  i2: 'Number',
  i4: 'Number',
  i8: 'Number',
  int: 'Number',
  ui1: 'Number',
  ui2: 'Number',
  ui4: 'Number',
  ui8: 'Number',
  uint: 'Number',
  r4: 'Number',
  r8: 'Number',
  decimal: 'Number',
  bool: 'Boolean',
  filetime: 'Date',
  date: 'Date',
});

/** One listed property: its element, its name, and its value. */
interface Listed {
  readonly el: Element;
  readonly read: AutomationCustomPropertyRead;
}

/** The value of a variant element as its listed type, or undefined when it does not parse. */
function valueOf(variant: Element): AutomationCustomPropertyRead['value'] | undefined {
  const type = TYPE_OF[variant.local];
  const text = textOf(variant);
  const raw = text.trim();
  if (type === 'String') return text;
  if (type === 'Number') {
    const n = raw === '' ? Number.NaN : Number(raw);
    return Number.isFinite(n) ? n : undefined;
  }
  if (type === 'Boolean') {
    const v = raw.toLowerCase();
    return v === 'true' || v === '1' ? true : v === 'false' || v === '0' ? false : undefined;
  }
  if (type === 'Date') {
    const ms = Date.parse(raw);
    return Number.isFinite(ms) ? new Date(ms).toISOString() : undefined;
  }
  return undefined;
}

/** The custom properties part, when the package has one with the right root. */
function customPart(ctx: PlanContext): string | undefined {
  const part = packagePart(ctx.batch.pkg, REL_CUSTOM);
  const root = rootOf(ctx.batch.pkg, part);
  if (!part || !root) return undefined;
  if (root.ns !== NS_CUSTOM || root.local !== 'Properties')
    throw new AutomationFailure(
      'unsupported-content',
      'the custom properties part has another root',
      'custom-properties-root'
    );
  return part;
}

/** The properties of a String, Number, Boolean, or Date type, in document order. */
function listed(root: Element | undefined): Listed[] {
  const out: Listed[] = [];
  for (const el of elements(root)) {
    if (el.ns !== NS_CUSTOM || el.local !== 'property') continue;
    const key = attr(el, 'name');
    const variant = elements(el).find((c) => c.ns === NS_VT);
    if (!key || !variant || TYPE_OF[variant.local] === undefined) continue;
    const value = valueOf(variant);
    if (value === undefined) continue;
    out.push({ el, read: { key, type: TYPE_OF[variant.local]!, value } });
  }
  return out;
}

/** Keys compare without case, as the object model looks them up. */
const sameKey = (a: string, b: string): boolean => a.toLowerCase() === b.toLowerCase();

export function readCustomProperties(ctx: PlanContext): AutomationValue {
  const root = rootOf(ctx.batch.pkg, customPart(ctx));
  return { kind: 'customProperties', properties: listed(root).map((p) => p.read) };
}

function keyOf(key: unknown): string {
  if (typeof key !== 'string' || key.length === 0 || key.length > MAX_CUSTOM_KEY)
    badArgument('key');
  if (!isValidXmlText(key)) badArgument('key');
  return key;
}

/** `ST_` text of `filetime` (22.4.2.13, an `xsd:dateTime` in UTC), whole seconds. */
function fileTime(iso: unknown): string {
  if (typeof iso !== 'string' || iso.length > 64) badArgument('value');
  const ms = Date.parse(iso);
  if (!Number.isFinite(ms)) badArgument('value');
  const year = new Date(ms).getUTCFullYear();
  // `FILETIME` counts from 1601; a year past 9999 does not print as an `xsd:dateTime`.
  if (year < 1601 || year > 9999) badArgument('value');
  return new Date(ms).toISOString().replace(/\.\d{3}Z$/, 'Z');
}

/** The variant element of a written value: `vt:lpwstr`, `vt:i4`, `vt:r8`, `vt:bool`, or `vt:filetime`. */
function variantSpec(value: AutomationCustomPropertyWrite): NodeSpec {
  if (typeof value !== 'object' || value === null) badArgument('value');
  let local: string;
  let content: string;
  switch (value.type) {
    case 'String':
      if (typeof value.value !== 'string' || value.value.length > MAX_CUSTOM_VALUE)
        badArgument('value');
      if (!isValidXmlText(value.value)) badArgument('value');
      [local, content] = ['lpwstr', value.value];
      break;
    case 'Number': {
      const n = value.value;
      if (typeof n !== 'number' || !Number.isFinite(n)) badArgument('value');
      const whole = Number.isInteger(n) && n >= -2_147_483_648 && n <= 2_147_483_647;
      [local, content] = whole ? ['i4', String(n + 0)] : ['r8', String(n)];
      break;
    }
    case 'Boolean':
      if (typeof value.value !== 'boolean') badArgument('value');
      [local, content] = ['bool', value.value ? 'true' : 'false'];
      break;
    case 'Date':
      [local, content] = ['filetime', fileTime(value.value)];
      break;
    default:
      badArgument('type');
  }
  return {
    ns: NS_VT,
    local,
    prefix: 'vt',
    attrs: [],
    children: content === '' ? [] : [{ text: content }],
  };
}

export function setCustomProperty(
  ctx: PlanContext,
  key: string,
  value: AutomationCustomPropertyWrite
): AutomationValue {
  const name = keyOf(key);
  const variant = variantSpec(value);
  customPart(ctx);
  const batch = ctx.batch;
  edit(() => {
    const part = ensurePart(batch, REL_CUSTOM, '/docProps/custom', CT_CUSTOM, CUSTOM_XML);
    const root = editableRoot(batch, part);
    const properties = elements(root).filter((e) => e.ns === NS_CUSTOM && e.local === 'property');
    // A property of any type with this key takes the new value: a key names one property.
    const existing = properties.find((e) => sameKey(attr(e, 'name') ?? '', name));
    if (existing) {
      for (const c of [...existing.children]) batch.push({ op: 'removeNode', part, node: c.id });
      batch.push({ op: 'insertNode', part, parent: existing.id, index: 0, spec: variant });
      return;
    }
    // `pid` values start at 2 and are unique in the part (22.3.2.2).
    let pid = 1;
    for (const p of properties) {
      const n = Number(attr(p, 'pid'));
      if (Number.isSafeInteger(n) && n > pid) pid = n;
    }
    if (pid >= 2_147_483_647)
      throw new AutomationFailure('invalid-argument', 'too many properties', 'key');
    batch.push({
      op: 'insertNode',
      part,
      parent: root.id,
      index: root.children.length,
      spec: {
        ns: NS_CUSTOM,
        local: 'property',
        attrs: [
          { ns: '', local: 'fmtid', value: FMTID },
          { ns: '', local: 'pid', value: String(pid + 1) },
          { ns: '', local: 'name', value: name },
        ],
        children: [variant],
      },
    });
  });
  return { kind: 'applied' };
}

export function deleteCustomProperty(ctx: PlanContext, key: string): AutomationValue {
  const name = keyOf(key);
  const part = customPart(ctx);
  const hit = listed(rootOf(ctx.batch.pkg, part)).find((p) => sameKey(p.read.key, name));
  if (!part || !hit) throw new AutomationFailure('item-not-found', 'no such property', 'key');
  const root = editableRoot(ctx.batch, part);
  const el = listed(root).find((p) => sameKey(p.read.key, name))!.el;
  ctx.batch.push({ op: 'removeNode', part, node: el.id });
  return { kind: 'applied' };
}

/** Delete every listed property. A property of another variant type stays. */
export function deleteCustomProperties(ctx: PlanContext): AutomationValue {
  const part = customPart(ctx);
  if (!part || listed(rootOf(ctx.batch.pkg, part)).length === 0) return { kind: 'applied' };
  for (const p of listed(editableRoot(ctx.batch, part)))
    ctx.batch.push({ op: 'removeNode', part, node: p.el.id });
  return { kind: 'applied' };
}
