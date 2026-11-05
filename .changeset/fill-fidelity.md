---
'@pptx-editor.dev/core': minor
---

Gradient, tiled-picture, and picture fills render closer to the reference PDFs (gradients gain `tileRect`), fill overlays draw (`ResolvedEffects.fillOverlay`, `BlendMode`), and `DocumentSession.imageSize` uses 72 DPI for an image without a stated resolution.
