// The own properties of the shared DOM prototypes, and the changes tests leave on them.
//
// `bun run test:dom` runs the React and the Vue adapter tests in one process, so every file
// shares one set of DOM classes. A test that "restores" a method with
// `HTMLElement.prototype.getBoundingClientRect = saved`, where happy-dom defines the method on
// `Element.prototype`, adds an own property to `HTMLElement.prototype`. That property hides
// every later patch of `Element.prototype`, in any later file. The preload (`preload-dom.ts`)
// fails the test that leaves such a copy, and the run that ends with any prototype change.

/** The DOM classes whose prototypes tests patch. */
export const GUARDED_CLASSES = [
  'EventTarget',
  'Node',
  'Element',
  'HTMLElement',
  'HTMLInputElement',
  'HTMLTextAreaElement',
  'HTMLDialogElement',
  'HTMLCanvasElement',
  'SVGElement',
  'Document',
] as const;

export type PrototypeSnapshot = ReadonlyMap<string, ReadonlyMap<PropertyKey, PropertyDescriptor>>;

function ownDescriptors(proto: object): Map<PropertyKey, PropertyDescriptor> {
  const out = new Map<PropertyKey, PropertyDescriptor>();
  for (const key of Reflect.ownKeys(proto)) {
    const d = Object.getOwnPropertyDescriptor(proto, key);
    if (d) out.set(key, d);
  }
  return out;
}

function prototypeOf(scope: object, name: string): object | undefined {
  const ctor = (scope as Record<string, unknown>)[name];
  return typeof ctor === 'function' ? (ctor.prototype as object) : undefined;
}

/** The descriptor that `key` resolves to above `proto` in its chain, or undefined. */
function inheritedDescriptor(proto: object, key: PropertyKey): PropertyDescriptor | undefined {
  for (let p = Object.getPrototypeOf(proto); p; p = Object.getPrototypeOf(p)) {
    const d = Object.getOwnPropertyDescriptor(p, key);
    if (d) return d;
  }
  return undefined;
}

const sameDescriptor = (a: PropertyDescriptor, b: PropertyDescriptor): boolean =>
  a.value === b.value && a.get === b.get && a.set === b.set;

/** The own property descriptors of each guarded prototype in `scope`. */
export function snapshotPrototypes(scope: object): PrototypeSnapshot {
  const out = new Map<string, Map<PropertyKey, PropertyDescriptor>>();
  for (const name of GUARDED_CLASSES) {
    const proto = prototypeOf(scope, name);
    if (proto) out.set(name, ownDescriptors(proto));
  }
  return out;
}

/**
 * Own properties added since `baseline` that only copy the inherited one: the trace of a
 * restore by assignment. No test needs one, and each hides later patches of the prototype it
 * copies from. Each one is deleted. Returns them as `Class.prototype.name`.
 */
export function shadowingCopies(scope: object, baseline: PrototypeSnapshot): string[] {
  const out: string[] = [];
  for (const [name, before] of baseline) {
    const proto = prototypeOf(scope, name);
    if (!proto) continue;
    for (const [key, d] of ownDescriptors(proto)) {
      if (before.has(key)) continue;
      const inherited = inheritedDescriptor(proto, key);
      if (!inherited || !sameDescriptor(d, inherited)) continue;
      out.push(`${name}.prototype.${String(key)}`);
      Reflect.deleteProperty(proto, key);
    }
  }
  return out;
}

/**
 * The properties that differ from `baseline`, as `Class.prototype.name`. Each one is put back:
 * an added property is deleted, and a changed or removed one gets its first descriptor again.
 */
export function prototypeChanges(scope: object, baseline: PrototypeSnapshot): string[] {
  const changes: string[] = [];
  for (const [name, before] of baseline) {
    const proto = prototypeOf(scope, name);
    if (!proto) continue;
    const now = ownDescriptors(proto);
    for (const [key, d] of now) {
      const old = before.get(key);
      if (old && sameDescriptor(old, d)) continue;
      changes.push(`${name}.prototype.${String(key)}`);
      if (old) Object.defineProperty(proto, key, old);
      else Reflect.deleteProperty(proto, key);
    }
    for (const [key, old] of before) {
      if (now.has(key)) continue;
      changes.push(`${name}.prototype.${String(key)}`);
      Object.defineProperty(proto, key, old);
    }
  }
  return changes;
}
