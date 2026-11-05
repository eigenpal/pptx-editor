---
'@pptx-editor.dev/pptx-to-pdf': minor
---

`openDocumentForExport` and `exportPdfFrom` export one deck more than once from one layout, `ExportResourceError` reports a disposed or aborted session, each PDF result maps its pages to deck slide numbers in `slideNumbers`, and an engine failure inside an export is reported as `PdfEncodingError`.
