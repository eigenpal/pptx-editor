---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/react': minor
'@pptx-editor.dev/i18n': minor
---

The core package now bundles its text shaper, so Bun and esbuild apps build, and `setHarfBuzzWasmUrl` points the shaper at a self-hosted `@pptx-editor.dev/core/harfbuzz.wasm`.
