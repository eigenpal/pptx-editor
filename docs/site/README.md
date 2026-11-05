# Site documentation source

This directory holds the public developer documentation in the DOCX Editor site structure (DOCX Editor `docs/site/README.md`, commit 68ad4957eeb9). A docs site syncs `content/` and `data/pptx-features.ts` at build time and renders the MDX pages. The site application is not in this repository. This file is for contributors; the site does not show it. Update these docs in the pull request that changes a user-facing feature.

## Layout

- `content/<slug>.mdx` is the page `/docs/<slug>`.
- `content/<dir>/index.mdx` is the page `/docs/<dir>`.
- `content/**/meta.json` sets the sidebar order and groups (the `pages` array; `---Label---` entries are group separators).
- `data/pptx-features.ts` is the feature matrix. The `<FeatureMatrix>`, `<FeatureSummary>`, and `<FeatureBadge>` components render it.

## Required frontmatter

```yaml
---
title: 'Installation' # Up to 60 characters; the site appends the brand.
description: 'Install the PPTX editor in React or Vue.' # Summarizes the page for search results.
category: 'Getting started' # Badge text and llms.txt group.
---
```

`seoTitle` (optional) is the long title for the HTML `<title>` and Open Graph tags. Keep `title` short: it is the page heading and the sidebar label.

## MDX components

The pages use these components without imports. A site that renders this directory must provide each of them.

| Component | Props | What it renders |
| --- | --- | --- |
| `FrameworkTabs`, `Framework` | `variant`, `value` | A React and Vue switch, as on the DOCX Editor site |
| `Cards`, `Card` | Fumadocs defaults | Link cards |
| `FeatureMatrix` | `category`: a key of `FEATURE_CATEGORY_LABELS` | One table of the rows of that category: name (linked to `docsLink` when set), editing, rendering, round trip, and notes |
| `FeatureSummary` | None | The number of rows, and the count for each editing status |
| `FeatureBadge` | `id`: a row `id` | The name and editing status of one row, inline |

The three feature components are the DOCX Editor site components (`src/components/docs/FeatureMatrix.tsx` in the DOCX Editor website repository), bound to `pptxFeatures`, `pptxFeatureById`, and `FEATURE_CATEGORY_LABELS` from `data/pptx-features.ts` in place of the `word-features.ts` exports. The data file imports nothing, so the site can copy it as it is. A row has only the fields of a DOCX Editor row. The components show no `tier`, so a premium row says in its notes that it needs the EigenPal Pro License, as in DOCX Editor. The evidence for each row is in `data/pptx-feature-evidence.ts`, which the site does not sync.

Update `data/pptx-features.ts` when the status of a feature changes. Never write support claims in prose. The notes show on the page, so they describe behavior only. `data/fidelity-page.test.ts` checks that each component tag names a category or a row that exists, and that each `docsLink` names a page.

## Conventions

Follow the [Google developer documentation style guide](https://developers.google.com/style/highlights). Lead with the task or behavior. Use the active voice, address the reader as "you," and use sentence case for headings. Keep paragraphs short. Format identifiers as code and UI labels in bold.

- Describe only what the published packages ship.
- Link between pages with root-relative paths: `[Fonts](/docs/customization/fonts)`.
- End each page with a short "See also" section.
- Register each page in `content/meta.json` and in its folder's `meta.json`.
- Introduce each code sample, and keep it consistent with the public API.
- Use Mermaid for diagrams. Add alt text to images and descriptive text to links.

## Validate changes

Run the docs checks from the repository root:

```bash
bun test docs
```

The checks validate the component tags, the matrix, and its evidence. They do not build or render the site.
