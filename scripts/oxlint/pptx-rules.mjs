// Repository lint rules that oxlint has no built-in for, in the shape of DOCX Editor's
// `scripts/oxlint/docx-rules.mjs`. ESLint expressed them as `no-restricted-syntax` selectors;
// oxlint has no such rule, so each concern is its own rule here, with the selectors and messages
// of the ESLint configuration this replaces. Adapted from DOCX Editor at commit adf0e178.
//
// One rule per concern also removes a trap of the ESLint configuration: there, a block that set
// `no-restricted-syntax` REPLACED the earlier value, so the lane block had to restate the sink
// selectors. Here a block turns on the rules it needs without touching the others.
//
// oxlint runs JavaScript plugins in alpha, outside semver. `scripts/oxlint/pptx-rules.test.ts`
// lints a fixture of every banned shape, so a version that stops running these rules fails.

/** Report every node a selector matches. */
function selectorRule(selectors) {
  return {
    create(context) {
      const visitors = {};
      for (const { selector, message } of selectors) {
        visitors[selector] = (node) => context.report({ node, message });
      }
      return visitors;
    },
  };
}

// AGENTS.md, "Security": every value from a `.pptx` is attacker-controlled, so a file-derived
// string must never reach an HTML parser.
const NO_HTML_SINK_MSG =
  'No HTML from strings: file-derived values must not reach an HTML parser. ' +
  'Use createElement(NS) + setAttribute/textContent. See AGENTS.md "Security".';

const NO_GLOBAL_KEYFRAMES_MSG =
  '@keyframes names are document-global. Give the name a pptx- prefix, or move the ' +
  'keyframes into packages/core/src/styles/editor.css.';

export default {
  meta: { name: 'pptx' },
  rules: {
    'no-html-sinks': selectorRule([
      {
        selector: "AssignmentExpression[left.property.name='innerHTML']",
        message: NO_HTML_SINK_MSG,
      },
      {
        selector: "AssignmentExpression[left.property.name='outerHTML']",
        message: NO_HTML_SINK_MSG,
      },
      {
        selector: "CallExpression[callee.property.name='insertAdjacentHTML']",
        message: NO_HTML_SINK_MSG,
      },
      {
        selector: "CallExpression[callee.object.name='document'][callee.property.name='write']",
        message: NO_HTML_SINK_MSG,
      },
      // `someWindow.document.write(...)`, the popup variant.
      {
        selector:
          "CallExpression[callee.object.property.name='document'][callee.property.name='write']",
        message: NO_HTML_SINK_MSG,
      },
      { selector: "JSXAttribute[name.name='dangerouslySetInnerHTML']", message: NO_HTML_SINK_MSG },
    ]),

    'no-global-keyframes': selectorRule([
      {
        selector: 'TemplateElement[value.raw=/@(-\\w+-)?keyframes\\s+(?!pptx-)/]',
        message: NO_GLOBAL_KEYFRAMES_MSG,
      },
      {
        selector: 'Literal[value=/@(-\\w+-)?keyframes\\s+(?!pptx-)/]',
        message: NO_GLOBAL_KEYFRAMES_MSG,
      },
    ]),

    // `array.push(...items)` passes every item as an argument. A file-derived array of 200k
    // items overflows the call stack. The DOM-free lanes handle untrusted sizes, so they loop.
    // Stricter than DOCX Editor, which grandfathers push, splice, and Math.* sites.
    'no-varargs-spread': selectorRule([
      {
        selector: 'CallExpression[callee.property.name=/^(push|unshift|splice)$/] > SpreadElement',
        message: 'Varargs spread is unbounded on file-derived arrays. Use a loop.',
      },
      {
        selector: "CallExpression[callee.object.name='Math'] > SpreadElement",
        message: 'Math.max/min(...items) is unbounded on file-derived arrays. Use a loop.',
      },
      {
        selector:
          "CallExpression:not([callee.object.name='Math']):not([callee.property.name=/^(push|unshift|splice)$/]) > SpreadElement",
        message:
          'Spreading an array into a call is unbounded on file-derived arrays. Pass the array.',
      },
      {
        selector: 'NewExpression > SpreadElement',
        message: 'Spreading an array into a constructor is unbounded. Pass the array itself.',
      },
    ]),
  },
};
