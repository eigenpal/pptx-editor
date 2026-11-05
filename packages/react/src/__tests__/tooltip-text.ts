// The tooltip a chrome control shows (`data-pptx-tooltip`), with a disabled control's reason on
// its own line, as the tooltip bubble draws them.

export function tooltipText(el: Element | null | undefined): string {
  if (!el) return '';
  return [el.getAttribute('data-pptx-tooltip'), el.getAttribute('data-pptx-tooltip-reason')]
    .filter(Boolean)
    .join('\n');
}
