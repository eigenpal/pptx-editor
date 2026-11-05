// One canonical form of a DOM tree for the cross-adapter tests: React and Vue render the same
// chrome when their canonical forms are equal, after DOCX Editor's cross-adapter differential
// test. Each test states the differences it ignores through the options, at its call site.

/** Attributes that hold generated ids or id references. */
const ID_ATTRIBUTES = new Set(['id', 'aria-describedby', 'aria-controls', 'aria-labelledby']);

export interface CanonicalOptions {
  /** Include the root element itself. By default only its children are walked. */
  readonly self?: boolean;
  /** Attribute names left out on every element. */
  readonly dropAttributes?: readonly string[];
  /**
   * Leave out the `value` attribute of a `select`. A select's value is its selected option, not
   * an attribute; happy-dom reflects the one that Vue sets.
   */
  readonly dropSelectValue?: boolean;
  /** Leave out an empty `class` attribute (React writes one on an SVG that Vue leaves out). */
  readonly dropEmptyClass?: boolean;
  /** Map each id in a space-separated id reference on its own, not the whole value. */
  readonly splitIdReferences?: boolean;
  /** Compare `style` without whitespace and a trailing semicolon. */
  readonly normalizeStyle?: boolean;
  /** Trim text nodes and leave out the ones that hold only whitespace. */
  readonly trimText?: boolean;
  /** Add the live form state (input value, checkbox state, select value) of each control. */
  readonly formState?: boolean;
  /** Elements whose children are left out, for example a painted slide. */
  readonly skipChildren?: (el: Element) => boolean;
}

/**
 * The canonical markup of `root`: attributes sorted, generated ids replaced by their order of
 * first use, and comments left out. A `null` root gives an empty string.
 */
export function canonicalDom(root: Element | null, options: CanonicalOptions = {}): string {
  if (!root) return '';
  const ids = new Map<string, string>();
  const idOf = (value: string): string => {
    if (!ids.has(value)) ids.set(value, `id${ids.size}`);
    return ids.get(value)!;
  };
  const idValue = (value: string): string =>
    options.splitIdReferences ? value.split(' ').map(idOf).join(' ') : idOf(value);
  const dropped = new Set(options.dropAttributes ?? []);

  const keep = (el: Element, name: string, value: string): boolean => {
    if (dropped.has(name)) return false;
    if (options.dropSelectValue && el.tagName === 'SELECT' && name === 'value') return false;
    if (options.dropEmptyClass && name === 'class' && value === '') return false;
    return true;
  };

  const attributeValue = (name: string, value: string): string => {
    if (ID_ATTRIBUTES.has(name)) return idValue(value);
    if (options.normalizeStyle && name === 'style')
      return value.replace(/\s+/g, '').replace(/;$/, '');
    return value;
  };

  const stateOf = (el: Element): string => {
    if (!options.formState) return '';
    if (el instanceof HTMLInputElement)
      return ` value=${JSON.stringify(el.type === 'checkbox' ? String(el.checked) : el.value)}`;
    if (el instanceof HTMLSelectElement) return ` selected=${JSON.stringify(el.value)}`;
    return '';
  };

  const walk = (node: Node): string => {
    if (node.nodeType === 3) {
      const text = node.textContent ?? '';
      return options.trimText ? text.trim() : text;
    }
    if (node.nodeType !== 1) return '';
    const el = node as Element;
    const attrs = [...el.attributes]
      .filter((a) => keep(el, a.name, a.value))
      .map((a) => `${a.name}="${attributeValue(a.name, a.value)}"`)
      .sort();
    const inner = options.skipChildren?.(el) ? '' : [...el.childNodes].map(walk).join('');
    const tag = el.tagName.toLowerCase();
    return `<${tag} ${attrs.join(' ')}${stateOf(el)}>${inner}</${tag}>`;
  };

  return options.self ? walk(root) : [...root.childNodes].map(walk).join('');
}
