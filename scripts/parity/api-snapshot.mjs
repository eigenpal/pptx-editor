// Read API Extractor snapshots (`docs/api/<slug>/index.api.md`) for the adapter parity checks.
// After DOCX Editor's scripts/lib/api-snapshot-parse.mjs and the `parseExtendsClause` /
// `effectiveInterfaceFields` pair in its scripts/check-parity-contract.mjs, reduced to what
// PPTX Editor checks. DOCX Editor commit adf0e178.

/** Every top-level exported name of a snapshot. */
export function exportedNames(snapshot) {
  const names = new Set();
  for (const line of snapshot.split(/\r?\n/)) {
    const declared =
      /^export (?:declare )?(?:abstract )?(?:function|const|let|interface|type|class|enum|namespace) ([A-Za-z0-9_$]+)/.exec(
        line
      );
    if (declared) {
      names.add(declared[1]);
      continue;
    }
    const reexport = /^export \{([^}]+)\}/.exec(line);
    if (reexport)
      for (const part of reexport[1].split(',')) {
        const name = part
          .trim()
          .split(/\s+as\s+/)
          .pop();
        if (name) names.add(name);
      }
  }
  return names;
}

/**
 * The release tag of each exported declaration: the `// @public`, `// @internal`, `// @beta`,
 * or `// @alpha` comment that API Extractor prints on the line before it. A name declared twice
 * (a function overload, a const with a type of the same name) keeps its first tag.
 */
export function releaseTags(snapshot) {
  const tags = new Map();
  let tag = null;
  for (const line of snapshot.split(/\r?\n/)) {
    const comment = /^\/\/ @(public|internal|beta|alpha)\b/.exec(line);
    if (comment) {
      tag = comment[1];
      continue;
    }
    const declared =
      /^export (?:declare )?(?:abstract )?(?:function|const|let|interface|type|class|enum|namespace) ([A-Za-z0-9_$]+)/.exec(
        line
      );
    if (declared && tag && !tags.has(declared[1])) tags.set(declared[1], tag);
    if (line.trim() !== '' && !line.startsWith('//')) tag = null;
  }
  return tags;
}

/** The lines of one `export interface <name>` block, header included, or null. */
function interfaceBlock(lines, name) {
  const start = lines.findIndex((l) => new RegExp(`^export interface ${name}\\b`).test(l));
  if (start === -1) return null;
  const header = [];
  let i = start;
  // The header can wrap; it ends at the line that opens the body.
  for (; i < lines.length; i += 1) {
    header.push(lines[i]);
    if (lines[i].trimEnd().endsWith('{')) break;
  }
  const body = [];
  for (i += 1; i < lines.length && !/^\}/.test(lines[i]); i += 1) body.push(lines[i]);
  return { header: header.join(' '), body };
}

/**
 * Parse `Base`, `Base<Args>`, and `Omit<Base, 'a' | 'b'>` terms from an extends clause. A
 * generic base keeps its name: the members are compared by name only.
 */
export function parseExtendsClause(clause) {
  const bases = [];
  let depth = 0;
  let start = 0;
  for (let i = 0; i <= clause.length; i += 1) {
    const c = clause[i];
    if (c === '<' || c === '(' || c === '{') depth += 1;
    else if (c === '>' || c === ')' || c === '}') depth -= 1;
    if ((c === ',' && depth === 0) || i === clause.length) {
      const part = clause.slice(start, i).trim();
      start = i + 1;
      if (!part) continue;
      const omit = /^Omit<([A-Za-z0-9_$.]+)(?:<.*>)?,\s*(.+)>$/s.exec(part);
      if (omit) {
        const omitted = new Set([...omit[2].matchAll(/'([^']+)'/g)].map((m) => m[1]));
        bases.push({ base: omit[1], omitted, text: part });
        continue;
      }
      const plain = /^([A-Za-z0-9_$.]+)(?:<.*>)?$/s.exec(part);
      bases.push({ base: plain ? plain[1] : part, omitted: new Set(), text: part });
    }
  }
  return bases;
}

function extendsClauseOf(header) {
  const m = /^export interface [A-Za-z0-9_$]+(?:<.*?>)?\s+extends\s+(.*)\{\s*$/.exec(header);
  return m ? m[1].trim() : null;
}

/**
 * The members of an exported interface, each with its normalized type text, with every base of
 * its `extends` clause folded in (an `Omit<Base, 'a'>` base minus its keys). Returns null when
 * the snapshot has no such interface. Throws when a base is not an interface of the same
 * snapshot: its members cannot be read, so the check would compare an incomplete set.
 */
export function interfaceMemberTypes(snapshot, name, seen = new Set()) {
  if (seen.has(name)) return new Map();
  seen.add(name);
  const lines = typeof snapshot === 'string' ? snapshot.split(/\r?\n/) : snapshot;
  const block = interfaceBlock(lines, name);
  if (!block) return null;
  const members = new Map();
  const clause = extendsClauseOf(block.header);
  if (clause)
    for (const { base, omitted, text } of parseExtendsClause(clause)) {
      const inherited = interfaceMemberTypes(lines, base, seen);
      if (!inherited)
        throw new Error(
          `${name} extends ${text}, which is not an interface of the snapshot; its members cannot be checked.`
        );
      for (const [member, type] of inherited) if (!omitted.has(member)) members.set(member, type);
    }
  for (const [member, type] of declaredMembers(block.body)) members.set(member, type);
  return members;
}

/** The member names of an exported interface, `extends` folded in (see interfaceMemberTypes). */
export function interfaceMembers(snapshot, name) {
  const types = interfaceMemberTypes(snapshot, name);
  return types ? new Set(types.keys()) : null;
}

/**
 * The members of an exported props type: an interface (see interfaceMemberTypes), or a type
 * alias built from interfaces, `Omit<…>`, and object literals with `&` and `|`, such as
 * `ToolbarSlotProps = ToolbarSlotBaseProps & ({ asChild: true; … } | { asChild?: false; … })`.
 * A member of every union arm keeps its type when the arms agree; otherwise its type is the
 * sorted union of the arms' types. Returns null when the snapshot has neither. Throws when a
 * named term is not a type of the snapshot.
 */
export function propsMemberTypes(snapshot, name) {
  const lines = snapshot.split(/\r?\n/);
  const asInterface = interfaceMemberTypes(lines, name);
  if (asInterface) return asInterface;
  const start = lines.findIndex((l) => new RegExp(`^export type ${name}\\b`).test(l));
  if (start === -1) return null;
  let text = lines[start].replace(/^export type [A-Za-z0-9_$]+(?:<.*?>)?\s*=\s*/, '');
  let depth = balance(text) + angleBalance(text);
  for (let i = start + 1; (depth > 0 || !/;\s*$/.test(text)) && i < lines.length; i += 1) {
    if (lines[i].trim().startsWith('//')) continue;
    text += ` ${lines[i].trim()}`;
    depth += balance(lines[i]) + angleBalance(lines[i]);
  }
  return termMembers(lines, text.replace(/;\s*$/, '').trim(), name);
}

function termMembers(lines, text, owner) {
  let term = text.trim();
  while (term.startsWith('(') && closing(term, 0) === term.length - 1)
    term = term.slice(1, -1).trim();
  const union = splitTopLevel(term, '|').filter((p) => p.trim());
  if (union.length > 1) {
    const arms = union.map((arm) => termMembers(lines, arm, owner));
    const out = new Map();
    for (const arm of arms)
      for (const [member, info] of arm) {
        const seen = out.get(member);
        out.set(
          member,
          seen
            ? {
                type: normalizeType(`${seen.type} | ${info.type}`),
                optional: seen.optional || info.optional,
              }
            : info
        );
      }
    return out;
  }
  const intersection = splitTopLevel(term, '&').filter((p) => p.trim());
  if (intersection.length > 1) {
    const out = new Map();
    for (const part of intersection)
      for (const [member, info] of termMembers(lines, part, owner)) out.set(member, info);
    return out;
  }
  if (term.startsWith('{')) {
    const out = new Map();
    for (const raw of splitTopLevel(term.slice(1, -1), ';')) {
      const m = /^\s*(?:readonly )?([A-Za-z0-9_$]+)(\??):\s*(.*?)\s*$/s.exec(raw);
      if (m) out.set(m[1], { type: normalizeType(m[3]), optional: m[2] === '?' });
    }
    return out;
  }
  const [base] = parseExtendsClause(term);
  const members = base && interfaceMemberTypes(lines, base.base);
  if (!members)
    throw new Error(
      `${owner} uses ${term}, which is not an interface of the snapshot; its members cannot be checked.`
    );
  for (const omitted of base.omitted) members.delete(omitted);
  return members;
}

function closing(text, open) {
  let depth = 0;
  for (let i = open; i < text.length; i += 1) {
    if (text[i] === '(') depth += 1;
    else if (text[i] === ')') {
      depth -= 1;
      if (depth === 0) return i;
    }
  }
  return -1;
}

/** Top-level members of an interface body: name → { type, optional }. */
function declaredMembers(body) {
  const out = new Map();
  for (let i = 0; i < body.length; i += 1) {
    const m = /^ {4}(?:readonly )?([A-Za-z0-9_$]+)(\??)([(:<])(.*)$/.exec(body[i]);
    if (!m) continue;
    let text = m[3] + m[4];
    let depth = balance(text);
    while (depth > 0 && i + 1 < body.length) {
      i += 1;
      text += ` ${body[i].trim()}`;
      depth += balance(body[i]);
    }
    text = text.replace(/;\s*$/, '');
    // `name(args): T` is the method form of `name: (args) => T`.
    const type = text.startsWith(':')
      ? text.slice(1)
      : text.replace(/^(<[^>]*>)?\((.*)\):\s*(.*)$/s, '$1($2) => $3');
    out.set(m[1], { type: normalizeType(type), optional: m[2] === '?' });
  }
  return out;
}

function balance(text) {
  let depth = 0;
  for (const c of text) {
    if (c === '{' || c === '(' || c === '[') depth += 1;
    else if (c === '}' || c === ')' || c === ']') depth -= 1;
  }
  return depth;
}

/**
 * One type text in a form both adapters can share: whitespace collapsed, no `readonly` on
 * members, no `undefined` member of a union, and the members of a top-level union sorted.
 */
export function normalizeType(text) {
  const flat = text
    .replace(/\s+/g, ' ')
    .replace(/\breadonly (?=[A-Za-z_$]+\??:)/g, '')
    .replace(/;\s*\}/g, ' }')
    .replace(/\s*;\s*/g, '; ')
    .trim();
  // The two snapshots quote string literals differently.
  const quoted = flat.replace(/"([^"'\\]*)"/g, "'$1'");
  const parts = [
    ...new Set(
      splitTopLevel(quoted, '|')
        .map((p) => p.trim())
        .filter((p) => p && p !== 'undefined')
    ),
  ];
  // `true | false` is `boolean`.
  if (parts.includes('true') && parts.includes('false')) {
    parts.splice(parts.indexOf('true'), 1);
    parts.splice(parts.indexOf('false'), 1, 'boolean');
  }
  return parts.length > 1 ? parts.sort().join(' | ') : (parts[0] ?? quoted);
}

function splitTopLevel(text, separator) {
  const parts = [];
  let depth = 0;
  let start = 0;
  for (let i = 0; i < text.length; i += 1) {
    const c = text[i];
    if ('<({['.includes(c)) depth += 1;
    else if ('>)}]'.includes(c) && !(c === '>' && text[i - 1] === '=')) depth -= 1;
    else if (c === separator && depth === 0) {
      parts.push(text.slice(start, i));
      start = i + 1;
    }
  }
  parts.push(text.slice(start));
  return parts;
}

const CONSTRUCTOR_TYPES = {
  StringConstructor: 'string',
  BooleanConstructor: 'boolean',
  NumberConstructor: 'number',
};

/**
 * The runtime props of an exported Vue component as the Vue snapshot prints them (the first
 * `ExtractPropTypes<{ … }>` block after `export const <name>`): name → normalized type text, or
 * null for a prop whose runtime type does not carry its TypeScript type (`Object`, `Array`,
 * `Function`, a constructor array). Returns null when the snapshot has no such block, for
 * example for a component typed through a namespace interface.
 */
export function vueRuntimePropTypes(snapshot, name) {
  const lines = snapshot.split(/\r?\n/);
  const start = lines.findIndex((l) => new RegExp(`^export const ${name}:`).test(l));
  if (start === -1) return null;
  let open = -1;
  for (let i = start; i < Math.min(lines.length, start + 3); i += 1)
    if (/ExtractPropTypes<\{$/.test(lines[i])) {
      open = i;
      break;
    }
  if (open === -1) return null;
  const indent = /^( *)/.exec(lines[open + 1] ?? '')[1].length;
  const out = new Map();
  let current = null;
  for (let i = open + 1; i < lines.length; i += 1) {
    const line = lines[i];
    const lead = /^( *)/.exec(line)[1].length;
    if (lead < indent) break;
    if (lead === indent) {
      const member = /^ *(?:readonly )?([A-Za-z0-9_$]+): (.*?);?$/.exec(line);
      if (!member) continue;
      current = member[1];
      out.set(current, member[2] === '{' ? null : runtimeType(member[2]));
      continue;
    }
    if (lead === indent + 4 && current) {
      const type = /^ *(?:readonly )?type: (.*)$/.exec(line);
      if (!type) continue;
      let text = type[1];
      let depth = balance(text) + angleBalance(text);
      while (depth > 0 && i + 1 < lines.length) {
        i += 1;
        text += ` ${lines[i].trim()}`;
        depth += balance(lines[i]) + angleBalance(lines[i]);
      }
      out.set(current, runtimeType(text.replace(/;\s*$/, '')));
    }
  }
  return out;
}

/** `<` minus `>`, not counting the `>` of an arrow. */
function angleBalance(text) {
  let depth = 0;
  for (let i = 0; i < text.length; i += 1) {
    if (text[i] === '<') depth += 1;
    else if (text[i] === '>' && text[i - 1] !== '=') depth -= 1;
  }
  return depth;
}

function runtimeType(text) {
  const propType = /^PropType<(.*)>$/s.exec(text.trim());
  if (propType) return normalizeType(propType[1]);
  return CONSTRUCTOR_TYPES[text.trim()] ?? null;
}

/**
 * Exported functions named `use*` with their declared return type name, for example
 * `{ name: 'useEditorState', returns: 'UseEditorStateReturn' }`.
 */
export function composables(snapshot) {
  const out = [];
  for (const line of snapshot.split(/\r?\n/)) {
    const m = /^export function (use[A-Z][A-Za-z0-9_]*)\b.*\): ([A-Za-z0-9_$]+)(?:<.*>)?;$/.exec(
      line
    );
    if (m) out.push({ name: m[1], returns: m[2] });
    else {
      const bare = /^export function (use[A-Z][A-Za-z0-9_]*)\b/.exec(line);
      if (bare) out.push({ name: bare[1], returns: null });
    }
  }
  return out;
}
