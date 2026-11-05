// The `store` lane: the bounded OPC/XML trust boundary, the canonical tree per part, the op
// vocabulary, and the transactional document store. DOM-free. The source of truth.

export * from './package/namespaces.ts';
export * from './package/tree.ts';
export * from './package/query.ts';
export * from './package/package.ts';
export * from './package/presentation.ts';
export * from './package/table-grid.ts';
export * from './package/table-marks.ts';
export { animatedShapeIds, timingRemovals } from './package/timing.ts';
export { serializeXmlPart, canonicalFingerprint, canonicalForm } from './package/serialize.ts';
export {
  sanitizeHref,
  escapeXml,
  escapeXmlAttribute,
  escapeCssString,
  isValidXmlText,
} from '@eigenpal/ooxml-core';
export { asciiFold, normalizePartName, resolveInternalTarget } from '@eigenpal/ooxml-core';
export { readZip, writeZip, DEFAULT_ZIP_LIMITS, type ZipLimits } from '@eigenpal/ooxml-core';
export * from './store/tree-edit.ts';
export * from './store/ops.ts';
export * from './store/document-store.ts';
export * from './store/history-group.ts';
export * from './store/replica-part.ts';
export {
  decodeEmbeddedFont,
  fontEmbeddingPermission,
  type DecodeEmbeddedFontResult,
  type EmbeddedFontInfo,
  type EmbeddingPermission,
} from './fonts/index.ts';
export { imageSize, imageSizeEmu, type ImageSize } from '@eigenpal/ooxml-core';
