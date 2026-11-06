# pptx-editor.dev

WYSIWYG editor and rendering engine for PPTX, based on ECMA-376.
One engine provides parsing, editing, layout, rendering, and save operations.
React and Vue provide adapters over that engine.

## Start here

1. Read [the repository map](docs/repo-map.md) for package locations and task-specific reading.
2. Read [the contribution guide](CONTRIBUTING.md) for setup, verification, and pull requests.
3. Read [the engine architecture](docs/architecture/engine-lanes.md) before changing package dependencies or data flow.
4. Read the relevant [maintained specifications](openspec/specs/) before changing behavior.

## Communication

Write replies in ASD-STE100 Simplified Technical English.
Use active voice, simple tenses, and one idea per sentence.
Keep sentences to 20 words and paragraphs to six sentences.
Keep technical identifiers exact.
Use `EigenPal Pro License` in user-facing text.
Preserve legal license files and license identifiers.

## Engineering rules

- Preserve unknown content through open, edit, save, and reopen.
- Write document changes only through `PptxDocumentStore` operations.
- Keep parsing, cascade, and layout free of DOM dependencies.
- Use the same layout and rendering pipeline in the editor and fidelity checks.
- Make the fidelity check compare two PDFs: the PDF of the editor's own paint and the reference PDF.
- Process both PDFs with the same tools. Do not add rules for an operating system, machine, or presentation application.
- Accept reference PDFs from any machine. Use only open fonts in fixtures.
- Read the [OOXML references](reference/README.md) before changing parsing, cascade, or layout.
- Measure rendering against reference PDFs. Never loosen thresholds to pass a check.
- Treat missing references and skipped checks as gaps, never passes.
- Keep React and Vue behavior aligned. Put shared logic in `core`.
- Keep `core` external to its adapters so consumers load one engine instance.
- Follow DOCX Editor's API patterns for shared concepts and customization.
- Preserve source attribution and license notices when sharing code.
- Verify collaboration convergence when changing editing behavior.
- Keep support claims in the [feature matrix](docs/site/data/pptx-features.ts), with corresponding evidence.

## Security and privacy

Treat every value from a presentation as untrusted input.
Validate at the parse boundary and escape values on save.
Keep XML, archive, font, and recursion limits enforced.
Use `sanitizeHref` for links and DOM APIs for rendered content.
Never execute embedded content or fetch external relationships automatically.
Respect font embedding permissions.
Never commit private presentations, their contents, or reference fonts.

## Names and documentation

Use ECMA-376 terminology for presentation features.
Do not use proprietary product comparisons or internal investigation details in committed text.
Preserve format identifiers, namespace URIs, package names, Office.js identifiers, font names, and required legal attribution.
Document shipped behavior in [the public documentation](docs/site/content/).
Keep [the repository map](docs/repo-map.md) current when modules move, appear, or disappear.

## Developer documentation

Write for developers who install, embed, customize, and automate the SDK.
Follow the [Google developer documentation style guide](https://developers.google.com/style) and its [highlights](https://developers.google.com/style/highlights).
Apply these rules to public guides, READMEs, API descriptions, feature descriptions, and release notes.

- Lead with the developer's task and the result. Give the smallest useful code example early.
- Use concise prose to explain decisions, prerequisites, defaults, lifecycle, and limits. Link to details instead of repeating them.
- Address the reader as "you." Use active voice, present tense, American spelling, and sentence-case headings.
- Put conditions before instructions. Use numbered lists for steps and parallel bullets for related items.
- Avoid marketing, idioms, and words such as "simply," "easy," "just," and "please" in procedures.
- Describe shipped behavior. Avoid time-dependent wording such as "currently," "soon," and "as of this writing."
- Format identifiers, filenames, and placeholders as code. Use bold for UI labels and descriptive text for links.
- Introduce every code sample. Include imports and required setup, or link to the setup and identify existing variables.
- Use comments for omitted code. Keep sample lines near 80 characters where practical.
- Show error handling and cleanup when the task requires them. Never leave resource ownership unclear.
- Check samples against package exports, public types, and implementation. Never infer an API from another SDK's names.
- For automation, load explicit properties, batch independent operations, and explain required `context.sync()` boundaries.
- Keep proxies inside their runtime scope. State browser-only and server-only behavior beside the example.
- Use proper heading order, image alt text, and accessible labels. Do not communicate meaning through color alone.
- Use Mermaid for useful diagrams. Do not use ASCII box diagrams or images of code.

Keep the DOCX Editor documentation structure for shared SDK concepts.
Use getting started, customization, frameworks, React, Vue, guides, export, Pro, editing API, and reference sections.
Use the same page names for equivalent tasks. Add presentation-specific pages where needed.
Do not create pages for capabilities that PPTX Editor does not ship.
Keep task guides separate from detailed API and component reference tables.
Show equivalent React and Vue samples with `FrameworkTabs`, React first, when a task applies to both adapters.
Use ordinary code fences on pages explicitly scoped to one adapter.

Author public pages in `docs/site/content/`. Follow [the site source conventions](docs/site/README.md).
Register every page in the root `meta.json` and its folder's `meta.json`.
Keep existing URLs when restructuring documentation, and check internal links after changes.
Keep feature-status claims in `docs/site/data/pptx-features.ts`; reference them with the site's feature components.
Explain API behavior and limits in guides without maintaining a second support matrix.
Update affected documentation in the same change as public behavior.
Run `bun test docs`, the public terminology check, and formatting checks for edited documentation.
Report source validation separately from a website build or browser review.

## Verification and release

Use [the contribution guide](CONTRIBUTING.md) and repository scripts for the checks relevant to each change.
Run the required checks before release and report failures or skipped checks accurately.
Build packages before checking API snapshots or generated notices.
Update API snapshots, feature evidence, and translations when the change affects them.
Add a changeset for code changes. Documentation-only changes do not need one.
Generate third-party notices from the built packages before publishing.
Keep required licenses and notices in each published package.
