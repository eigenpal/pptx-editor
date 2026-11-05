#!/usr/bin/env bun
// Cross-adapter parity between @pptx-editor.dev/react and @pptx-editor.dev/vue, after DOCX
// Editor's scripts/check-parity-contract.mjs. The React side comes from its API Extractor
// snapshot (`docs/api/react/index.api.md`), with each props interface's `extends` clause folded
// in. The Vue side comes from the components themselves: their runtime props and emits, and
// mount probes for `class`, `className`, and every mapped slot
// (`packages/vue/test/parity-shape.ts`). The Vue prop types derive from those runtime props, so
// a hand-written interface cannot hide a drift. The contract is
// `scripts/parity/parity.contract.json`.
//
// Fails on any drift the contract does not acknowledge:
// - a React component export with a props type has no `components` entry and no
//   `propsUnchecked` reason;
// - a prop, emit, or export exists in one adapter but the contract does not classify it;
// - a `paired` entry is missing on one side, or its type differs between the two snapshots;
// - a `deferredInVue` entry has shipped in Vue, or is gone from React (stale);
// - a `vueExclusive` entry appeared in React, or is gone from Vue (stale);
// - a React callback, render prop, or className has no working Vue emit, slot, or class;
// - a value probe shows that a Vue part ignores a paired prop;
// - a compound sub-part (`"partOf": "PptxEditorToolbar.FontFamily"`, reached as a static, never
//   exported) is gated like a component, but its namespace export must exist in both adapters
//   instead of the part itself;
// - a `PptxEditorRef` member exists in one adapter but the contract's `ref` section does not
//   classify it, or a paired member is missing on one side or has a different type;
// - a paired or renamed export has a different release tag (`@public`, `@internal`) in the two
//   snapshots, so one adapter cannot announce a building block that the other keeps internal.
//
// The Vue parts declare `className` and `style` props and take host-owned callbacks as
// function props with React's names, so those are `paired`. Only an entry with
// `"topLevel": true` (the `PptxEditor` sugar, as DOCX Editor's `DocxEditor`) must take its
// callbacks as emits and its `className` as a fallthrough `class`, never as props.
//
// Run with Bun: it imports the Vue adapter's TypeScript sources.

import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  exportedNames,
  interfaceMemberTypes,
  propsMemberTypes,
  releaseTags,
  vueRuntimePropTypes,
} from './parity/api-snapshot.mjs';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');

const asList = (value) => (Array.isArray(value) ? value : [value]);

/**
 * Prop types are compared only where both snapshots print the TypeScript type: React's
 * interface member, and a Vue runtime prop declared with `PropType<T>`, `String`, `Boolean`, or
 * `Number`. A Vue prop declared as `Object`, `Array`, `Function`, or a constructor array carries
 * no TypeScript type, and a Vue component typed through a namespace interface prints no runtime
 * props; those pairs are checked by name only. `style` is skipped: React's `CSSProperties` and
 * Vue's `CSSProperties` are two framework types with one name.
 */
const TYPE_UNCHECKED = new Set(['style']);

/**
 * A rendered node is a framework type: React's `ReactNode` against Vue's `VNode` or
 * `VNodeChild`. Both name "anything the part renders", strings and numbers included, so a union
 * that holds one compares as `Node` with its primitive members dropped.
 */
function frameworkNeutral(type) {
  const mapped = type.replace(
    /\b(?:react\.|vue\.)?(?:ReactNode|ReactElement|JSX\.Element|VNodeChild|VNode)\b/g,
    'Node'
  );
  // `ReactElement<Props>` names a rendered node too: its props are React's own typing.
  const parts = mapped.replace(/\bNode<[^<>]*>/g, 'Node').split(' | ');
  if (!parts.includes('Node')) return mapped;
  const kept = parts.filter((p) => !['string', 'number', 'boolean', 'null'].includes(p));
  return [...new Set(kept)].sort().join(' | ');
}

function componentErrors(name, rules, { reactApi, vueApi }, vue) {
  const errors = [];
  let reactTypes;
  try {
    reactTypes = propsMemberTypes(reactApi, rules.reactProps);
  } catch (error) {
    return [`${name} (${rules.reactProps}): ${error.message}`];
  }
  if (!reactTypes) return [`${name}: React snapshot has no ${rules.reactProps} type.`];
  if (!vue) return [`${name}: the Vue adapter has no ${name} component, or it has no probe.`];
  const reactProps = new Set(reactTypes.keys());
  const vueProps = new Set(vue.props);
  const vueEmits = new Set(vue.emits);
  const vueSlots = new Set(vue.slots);
  const where = `${name} (${rules.reactProps})`;
  const emitsMap = rules.reactCallbacksAsVueEmits;
  const slotsMap = rules.reactRenderPropsAsVueSlots;
  const classMap = rules.reactClassNameAsVueClass;
  const reactClassified = new Set([
    ...rules.paired,
    ...Object.keys(rules.deferredInVue),
    ...Object.keys(emitsMap),
    ...Object.keys(slotsMap),
    ...Object.keys(classMap),
  ]);
  const vueClassified = new Set([...rules.paired, ...Object.keys(rules.vueExclusive)]);

  for (const prop of reactProps)
    if (!reactClassified.has(prop))
      errors.push(`${where}: React prop "${prop}" is not classified in the contract.`);
  for (const prop of vueProps)
    if (!vueClassified.has(prop))
      errors.push(`${where}: Vue prop "${prop}" is not classified in the contract.`);

  const vueTypes = vueRuntimePropTypes(vueApi, name);
  for (const prop of rules.paired) {
    if (!reactProps.has(prop)) errors.push(`${where}: paired "${prop}" is missing in React.`);
    if (!vueProps.has(prop)) errors.push(`${where}: paired "${prop}" is missing in Vue.`);
    const reactType = reactTypes.get(prop)?.type;
    const vueType = vueTypes?.get(prop);
    const react = reactType && frameworkNeutral(reactType);
    if (!TYPE_UNCHECKED.has(prop) && react && vueType && react !== frameworkNeutral(vueType))
      errors.push(`${where}: paired "${prop}" type drift: React \`${react}\`, Vue \`${vueType}\`.`);
  }
  for (const prop of Object.keys(rules.deferredInVue)) {
    if (!reactProps.has(prop))
      errors.push(`${where}: deferredInVue "${prop}" is missing in React (contract stale).`);
    if (vueProps.has(prop))
      errors.push(`${where}: deferredInVue "${prop}" has shipped in Vue; move it to paired.`);
  }
  for (const prop of Object.keys(rules.vueExclusive)) {
    if (!vueProps.has(prop))
      errors.push(`${where}: vueExclusive "${prop}" is missing in Vue (contract stale).`);
    if (reactProps.has(prop))
      errors.push(`${where}: vueExclusive "${prop}" appeared in React; reclassify it.`);
  }

  const declaredEmits = new Set();
  for (const [react, emits] of Object.entries(emitsMap)) {
    if (!reactProps.has(react)) errors.push(`${where}: React has no "${react}" callback.`);
    if (rules.topLevel && vueProps.has(react))
      errors.push(`${where}: "${react}" must be a Vue emit, not a prop.`);
    for (const emit of asList(emits)) {
      declaredEmits.add(emit);
      if (!vueEmits.has(emit)) errors.push(`${where}: Vue ${name} does not emit "${emit}".`);
    }
  }
  for (const emit of vueEmits)
    if (!declaredEmits.has(emit))
      errors.push(`${where}: Vue emit "${emit}" is not declared in reactCallbacksAsVueEmits.`);

  for (const [react, slot] of Object.entries(slotsMap)) {
    if (!reactProps.has(react)) errors.push(`${where}: React has no "${react}" render prop.`);
    // A part may take the render function as a prop too (DOCX Editor's open-failure `actions`):
    // then it is also `paired`, and the slot wins over the prop.
    if (vueProps.has(react) && !rules.paired.includes(react))
      errors.push(`${where}: "${react}" must be a Vue slot, not a prop.`);
    if (!vueSlots.has(slot))
      errors.push(`${where}: Vue ${name} does not render its "${slot}" slot.`);
  }

  for (const react of Object.keys(classMap)) {
    if (!reactProps.has(react)) errors.push(`${where}: React has no "${react}" prop.`);
    if (rules.topLevel && vueProps.has(react))
      errors.push(`${where}: "${react}" must fall through as a Vue class, not be a prop.`);
  }
  // A host class lands wherever the part takes one: as the fallthrough `class` of the top-level
  // component, or merged with the `className` prop of a part. `classUnprobed` names a part the
  // probe cannot make render (a notice with no state to show) and the test that checks it
  // instead; `hostClassGap` names a part whose host `class` is a known gap. Both go stale, and
  // fail, once the probe sees the class.
  const takesClass = Object.keys(classMap).length > 0 || rules.paired.includes('className');
  const unprobed = typeof rules.classUnprobed === 'string';
  const gap = typeof rules.hostClassGap === 'string';
  if (takesClass && !unprobed && !gap && !vue.classLands)
    errors.push(`${where}: a class on Vue ${name} does not reach a rendered element.`);
  if (rules.paired.includes('className') && !unprobed && !vue.classNameLands)
    errors.push(`${where}: the className prop of Vue ${name} does not reach a rendered element.`);
  if (unprobed && vue.classLands && vue.classNameLands)
    errors.push(`${where}: the probe now sees the class; remove classUnprobed.`);
  if (gap && vue.classLands)
    errors.push(`${where}: a host class now reaches an element; remove hostClassGap.`);
  // Value probes: a paired prop given a value shows in the rendered Vue part.
  for (const [prop, shows] of Object.entries(vue.values ?? {})) {
    if (!rules.paired.includes(prop))
      errors.push(`${where}: the value probe for "${prop}" names a prop that is not paired.`);
    if (!shows) errors.push(`${where}: Vue ${name} ignores a value of its paired "${prop}" prop.`);
  }
  return errors;
}

/** PascalCase React exports that are components with a props parameter. */
function reactComponentExports(reactApi) {
  const out = new Set();
  for (const line of reactApi.split(/\r?\n/)) {
    const fn = /^export function ([A-Z][A-Za-z0-9_$]*)\((?!\))/.exec(line);
    if (fn) out.add(fn[1]);
    const constant = /^export const ([A-Z][A-Za-z0-9_$]*): (.*)$/.exec(line);
    if (constant && !/^react\.Context</.test(constant[2])) out.add(constant[1]);
  }
  return out;
}

function coverageErrors({ reactApi, vueApi }, contract) {
  const errors = [];
  const renamed = contract.exports.reactNamesAsVueNames;
  const vueToReact = new Map(Object.entries(renamed).map(([react, vue]) => [vue, react]));
  const unchecked = contract.propsUnchecked ?? {};
  const covered = new Map();
  const vueExports = exportedNames(vueApi);
  const reactExports = exportedNames(reactApi);
  for (const [name, rules] of Object.entries(contract.components)) {
    if (name === 'comment') continue;
    // A compound sub-part is a static of a namespace both adapters export: React reaches it as
    // `PptxEditor.Toolbar.FontFamily.Trigger`, Vue as `PptxEditorToolbar.FontFamily.Trigger`.
    if (typeof rules.partOf === 'string') {
      const root = rules.partOf.split('.')[0];
      if (!vueExports.has(root)) errors.push(`components: Vue exports no "${root}" for ${name}.`);
      if (!reactExports.has('PptxEditor'))
        errors.push(`components: React exports no PptxEditor namespace for ${name}.`);
      continue;
    }
    const react = rules.reactComponent ?? vueToReact.get(name) ?? name;
    covered.set(react, name);
    if (!vueExports.has(name)) errors.push(`components: Vue exports no "${name}".`);
  }
  const components = reactComponentExports(reactApi);
  for (const react of covered.keys())
    if (!components.has(react))
      errors.push(`components: React exports no component "${react}" (contract stale).`);
  for (const react of components)
    if (!covered.has(react) && !(react in unchecked))
      errors.push(
        `components: React component "${react}" has no components entry and no propsUnchecked reason.`
      );
  for (const [react, reason] of Object.entries(unchecked)) {
    if (react === 'comment') continue;
    if (!components.has(react))
      errors.push(`propsUnchecked: React exports no component "${react}" (contract stale).`);
    if (covered.has(react))
      errors.push(`propsUnchecked: "${react}" has a components entry; remove it here.`);
    if (typeof reason !== 'string' || !reason.trim())
      errors.push(`propsUnchecked: "${react}" needs a reason.`);
  }
  return errors;
}

/**
 * The ref gate, after DOCX Editor's `DocxEditorRef` check: each adapter declares an
 * `interface PptxEditorRef` of its own, and every member of either one is `paired` (same name
 * and type on both), `deferredInVue`, or `vueExclusive` in the contract's `ref` section.
 */
function refErrors({ reactApi, vueApi }, rules) {
  const errors = [];
  const react = interfaceMemberTypes(reactApi, 'PptxEditorRef');
  const vue = interfaceMemberTypes(vueApi, 'PptxEditorRef');
  if (!react) errors.push('ref: React snapshot has no PptxEditorRef interface.');
  if (!vue) errors.push('ref: Vue snapshot has no PptxEditorRef interface.');
  if (!react || !vue) return errors;
  const paired = rules.paired ?? [];
  const deferred = Object.keys(rules.deferredInVue ?? {});
  const vueOnly = Object.keys(rules.vueExclusive ?? {});
  for (const name of react.keys())
    if (!paired.includes(name) && !deferred.includes(name))
      errors.push(`ref: React PptxEditorRef member "${name}" is not classified in the contract.`);
  for (const name of vue.keys())
    if (!paired.includes(name) && !vueOnly.includes(name))
      errors.push(`ref: Vue PptxEditorRef member "${name}" is not classified in the contract.`);
  for (const name of paired) {
    const r = react.get(name);
    const v = vue.get(name);
    if (!r) errors.push(`ref: paired "${name}" is missing in React.`);
    if (!v) errors.push(`ref: paired "${name}" is missing in Vue.`);
    if (r && v && (r.type !== v.type || r.optional !== v.optional))
      errors.push(`ref: paired "${name}" type drift: React \`${r.type}\`, Vue \`${v.type}\`.`);
  }
  for (const name of deferred) {
    if (!react.has(name)) errors.push(`ref: deferredInVue "${name}" is missing in React (stale).`);
    if (vue.has(name)) errors.push(`ref: deferredInVue "${name}" has shipped in Vue; pair it.`);
  }
  for (const name of vueOnly) {
    if (!vue.has(name)) errors.push(`ref: vueExclusive "${name}" is missing in Vue (stale).`);
    if (react.has(name)) errors.push(`ref: vueExclusive "${name}" appeared in React; pair it.`);
  }
  return errors;
}

/**
 * Every parity error for the React snapshot, the Vue snapshot, the Vue runtime shape
 * (`vueAdapterShape()`), and the contract.
 */
export function parityErrors({ reactApi, vueApi, vue, contract }) {
  const errors = [];
  for (const [name, rules] of Object.entries(contract.components)) {
    if (name === 'comment') continue;
    errors.push(...componentErrors(name, rules, { reactApi, vueApi }, vue[name]));
  }
  errors.push(...coverageErrors({ reactApi, vueApi }, contract));
  errors.push(...refErrors({ reactApi, vueApi }, contract.ref ?? {}));

  const reactExports = exportedNames(reactApi);
  const vueExports = exportedNames(vueApi);
  const e = contract.exports;
  const renamed = e.reactNamesAsVueNames;
  const reactClassified = new Set([
    ...e.paired,
    ...Object.keys(renamed),
    ...Object.keys(e.deferredInVue),
    ...Object.keys(e.reactExclusive),
  ]);
  const vueClassified = new Set([
    ...e.paired,
    ...Object.values(renamed),
    ...Object.keys(e.vueExclusive),
  ]);
  for (const name of reactExports)
    if (!reactClassified.has(name))
      errors.push(`exports: React export "${name}" is not classified in the contract.`);
  for (const name of vueExports)
    if (!vueClassified.has(name))
      errors.push(`exports: Vue export "${name}" is not classified in the contract.`);
  for (const name of e.paired) {
    if (!reactExports.has(name)) errors.push(`exports: paired "${name}" is missing in React.`);
    if (!vueExports.has(name)) errors.push(`exports: paired "${name}" is missing in Vue.`);
  }
  for (const [react, vueName] of Object.entries(renamed)) {
    if (!reactExports.has(react)) errors.push(`exports: React has no "${react}" (renamed pair).`);
    if (!vueExports.has(vueName))
      errors.push(`exports: Vue has no "${vueName}" (pair of "${react}").`);
  }
  // The two adapters give a pair the same release tag.
  const reactTags = releaseTags(reactApi);
  const vueTags = releaseTags(vueApi);
  const pairs = [...e.paired.map((name) => [name, name]), ...Object.entries(renamed)];
  for (const [react, vueName] of pairs) {
    const reactTag = reactTags.get(react);
    const vueTag = vueTags.get(vueName);
    if (reactTag && vueTag && reactTag !== vueTag)
      errors.push(
        `exports: "${react}" is @${reactTag} in React but "${vueName}" is @${vueTag} in Vue.`
      );
  }
  for (const name of Object.keys(e.deferredInVue)) {
    if (!reactExports.has(name))
      errors.push(`exports: deferredInVue "${name}" is missing in React (contract stale).`);
    if (vueExports.has(name))
      errors.push(`exports: deferredInVue "${name}" has shipped in Vue; move it to paired.`);
  }
  for (const name of Object.keys(e.reactExclusive)) {
    if (!reactExports.has(name))
      errors.push(`exports: reactExclusive "${name}" is missing in React (contract stale).`);
    if (vueExports.has(name))
      errors.push(`exports: reactExclusive "${name}" appeared in Vue; reclassify it.`);
  }
  for (const name of Object.keys(e.vueExclusive)) {
    if (!vueExports.has(name))
      errors.push(`exports: vueExclusive "${name}" is missing in Vue (contract stale).`);
    if (reactExports.has(name))
      errors.push(`exports: vueExclusive "${name}" appeared in React; reclassify it.`);
  }
  return errors;
}

async function main() {
  const read = (path) => readFileSync(join(root, path), 'utf8');
  const contract = JSON.parse(read('scripts/parity/parity.contract.json'));
  const { vueAdapterShape } = await import('../packages/vue/test/parity-shape.ts');
  const errors = parityErrors({
    reactApi: read('docs/api/react/index.api.md'),
    vueApi: read('docs/api/vue/index.api.md'),
    vue: await vueAdapterShape(contract),
    contract,
  });
  if (errors.length) {
    console.error('✘ adapter parity contract:\n');
    for (const error of errors) console.error(`  - ${error}`);
    console.error('\nUpdate the adapter or scripts/parity/parity.contract.json.');
    process.exit(1);
  }
  console.log('✓ React and Vue adapters match scripts/parity/parity.contract.json.');
  process.exit(0);
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) await main();
