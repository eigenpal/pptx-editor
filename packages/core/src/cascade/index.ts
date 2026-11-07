// The `cascade` lane: master → layout → slide inheritance, theme colors and fonts, style
// references, and list styles. DOM-free. Layout and paint read only what this lane resolves.

export * from './types.ts';
export * from './color.ts';
export * from './theme.ts';
export * from './paint-props.ts';
export * from './text-style.ts';
export { inheritedPlaceholders, layoutPlaceholderOf, resolveSlide, XFRM_BOUND } from './slide.ts';
export { CHART_CONTENT_TYPE, resolveChartPart, type ResolvedChartPart } from './chart-part.ts';
export { alternateBranch, contentElements, isAlternateContent } from './alternate-content.ts';
export { resolveTable, type TableScope } from './table.ts';
export type { PictureBulletSource } from './picture-bullet.ts';
export {
  DEFAULT_TABLE_STYLE_ID,
  defaultTableStyleId,
  tableStyleFor,
  tableStyleList,
} from './table-styles.ts';
export { PRESET_COLORS } from '@eigenpal/ooxml-core';
