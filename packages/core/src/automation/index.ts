// The automation lane: DOM-free edits that lower to store ops (`commands/`), and the automation
// host that runs Office.js-shaped batches over them, after DOCX Editor's
// `packages/core/src/automation`. `@pptx-editor.dev/editor-api` is the object model over this
// protocol. DOCX Editor commit adf0e178.

export type { LineEndType, ShapeFormatPatch, ShapeFormatProperties } from '../contracts/neutral.ts';
export type {
  AutomationAdjustmentsRead,
  AutomationBackgroundFillRead,
  AutomationBackgroundRead,
  AutomationBatchRequest,
  AutomationBatchResponse,
  AutomationBindingRead,
  AutomationCapabilities,
  AutomationChangeEvent,
  AutomationCustomPropertyRead,
  AutomationCustomXmlPartRead,
  AutomationError,
  AutomationErrorCode,
  AutomationFillRead,
  AutomationFontRead,
  AutomationHandle,
  AutomationHandleRef,
  AutomationHost,
  AutomationHyperlinkRead,
  AutomationImageRequest,
  AutomationImageResult,
  AutomationLayoutRead,
  AutomationLineRead,
  AutomationObjectKind,
  AutomationOperationResult,
  AutomationParagraphRead,
  AutomationPlaceholderRead,
  AutomationPresentationRead,
  AutomationPropertiesRead,
  AutomationRef,
  AutomationResultRef,
  AutomationSaveResult,
  AutomationSelectionRead,
  AutomationShapeRead,
  AutomationSlideRead,
  AutomationTableCellFormatRead,
  AutomationTableCellRead,
  AutomationTableRead,
  AutomationTag,
  AutomationTextFrameRead,
  AutomationTextRunRead,
  AutomationUnsubscribe,
  AutomationValue,
} from './protocol.ts';
export type {
  AutomationBackgroundWrite,
  AutomationBox,
  AutomationCellAddress,
  AutomationCellBorderSide,
  AutomationCellMarginsWrite,
  AutomationCustomPropertyWrite,
  AutomationFillWrite,
  AutomationFontWrite,
  AutomationInsertSlidesOptions,
  AutomationItemKey,
  AutomationLineWrite,
  AutomationNewShape,
  AutomationOperation,
  AutomationParagraphWrite,
  AutomationPropertiesWrite,
  AutomationShapeWrite,
  AutomationTableStyleWrite,
  AutomationTextFrameWrite,
  AutomationTextSpan,
  AutomationZOrder,
} from './operations.ts';
export { AUTOMATION_COMMANDS } from './operations.ts';
export {
  createAutomationHost,
  MAX_BATCH_OPERATIONS,
  type AutomationDocumentPort,
  type AutomationHostComposition,
} from './host.ts';
export type { AutomationSelectionPort } from './plan.ts';
export type { AutomationImageTarget } from './plan-images.ts';
export {
  createServerAutomationHost,
  createStoreAutomationHost,
  storeAutomationPort,
  type ServerAutomationOptions,
  type ServerAutomationRejection,
  type StoreAutomationOptions,
} from './server-host.ts';
