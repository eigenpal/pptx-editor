# pptx-editor.dev

WYSIWYG editor and rendering engine for PPTX, based on ECMA-376.
One engine parses, edits, lays out, renders, and saves. React and Vue are adapters over it.

## Start here

1. [Repository map](docs/repo-map.md): package locations and task-specific reading.
2. [Contribution guide](CONTRIBUTING.md): setup, verification, and pull requests.
3. [Engine architecture](docs/architecture/engine-lanes.md): read before changing dependencies or data flow.
4. [Maintained specifications](openspec/specs/): read before changing behavior.

## Communication

Write replies in ASD-STE100 Simplified Technical English.
Use active voice, simple tenses, one idea per sentence, and 20 words or fewer per sentence.
Keep technical identifiers exact.
Use `EigenPal Pro License` in user-facing text. Preserve legal license files and license identifiers.

## Engineering rules

- Preserve unknown content through open, edit, save, and reopen.
- Write document changes only through `PptxDocumentStore` operations.
- Keep parsing, cascade, and layout free of DOM dependencies.
- Use one layout and rendering pipeline for the editor and the fidelity checks.
- Compare two PDFs in the fidelity check: the editor's own paint and the reference PDF.
- Process both PDFs with the same tools. Accept reference PDFs from any machine.
- Use only open fonts in fixtures.
- Never loosen thresholds to pass a check. Treat missing references and skipped checks as gaps.
- Read the [OOXML references](reference/README.md) before changing parsing, cascade, or layout.
- Keep React and Vue behavior aligned. Put shared logic in `core`, and keep `core` external to its adapters.
- Follow DOCX Editor's API patterns for shared concepts and customization.
- State the source of every measured or adopted value. Preserve attribution and license notices.
- Verify collaboration convergence when changing editing behavior.
- Keep support claims in the [feature matrix](docs/site/data/pptx-features.ts), with evidence.

## Security and privacy

Treat every value from a presentation as untrusted input.
Validate at the parse boundary and escape values on save.
Keep XML, archive, font, and recursion limits enforced.
Use `sanitizeHref` for links and DOM APIs for rendered content.
Never execute embedded content or fetch external relationships automatically.
Respect font embedding permissions.
Never commit private presentations, their contents, or reference fonts.

## Developer documentation

Follow the [Google developer documentation style guide](https://developers.google.com/style/highlights) and the [site source conventions](docs/site/README.md).

- Author public pages in `docs/site/content/`. Register each page in the root `meta.json` and its folder's `meta.json`.
- Lead with the developer's task and a small code example. Describe shipped behavior only.
- Check samples against package exports and public types. State browser-only and server-only behavior.
- Keep the DOCX Editor page structure for shared SDK concepts. Keep existing URLs.
- Keep feature status in `docs/site/data/pptx-features.ts`, not in prose.
- Update documentation and [the repository map](docs/repo-map.md) in the same change as the behavior.
- Run `bun test docs` and Prettier on edited documentation.

## Verification and release

Run the checks in [the contribution guide](CONTRIBUTING.md) that apply to the change. Report failures and skipped checks accurately.
Build packages before checking API snapshots or generated notices.
Update API snapshots, feature evidence, and translations when the change affects them.
Add a changeset for code changes. Documentation-only changes do not need one.
Keep required licenses and notices in each published package.

## Naming

Use ECMA-376 terminology for presentation features.
Preserve format identifiers, namespace URIs, package names, Office.js identifiers, font names, and required legal attribution.

Trademark policy: do not use third-party product names in public text.
