// Specialized objects reuse the synthetic fixtures and their complete package relationships.
// These slides keep editable chart data, diagram models, and Office Math in the demo.

import type { FixtureOptions, FixtureSlide } from '../fixtures/builder.ts';
import { chartDataDeck } from '../fixtures/chart-data.ts';
import { diagramEditingDeck } from '../fixtures/diagram-editing.ts';
import { equationFixtures } from '../fixtures/equations.ts';

/** Three additional demo slides and the parts their objects require. */
export function specializedShowcase(): FixtureOptions {
  const chart = chartDataDeck();
  const diagram = diagramEditingDeck();
  const slides: FixtureSlide[] = [
    {
      ...chart.slides[0]!,
      notes: [
        'This chart keeps its values in an embedded workbook.',
        'Try: select the chart and open its formatting controls in the right panel.',
        'Try: edit a category or series value, then undo the edit.',
        'Try: change the legend position or a series color.',
      ].join('\n'),
    },
    {
      ...diagram.slides[0]!,
      notes: [
        'This process diagram has an editable data model and a saved drawing.',
        'Try: select the diagram and open its text controls in the right panel.',
        'Try: change a node label, then add a node and undo the addition.',
      ].join('\n'),
    },
    {
      ...equationFixtures.equations![0]!,
      // Simple scripts use the demo's available open serif face. The reference fixture stays unchanged.
      shapes: equationFixtures.equations![0]!.shapes.replaceAll(
        'STIX Two Math',
        'Liberation Serif'
      ),
      notes: [
        'These text boxes contain Office Math with plain-text fallback content.',
        'Try: select an equation and open the equation editor.',
        'Try: change a supported expression, then save and reopen the deck.',
      ].join('\n'),
    },
  ];
  return {
    slides,
    extraParts: [...(chart.extraParts ?? []), ...(diagram.extraParts ?? [])].filter(([name]) =>
      /1\.xml(?:\.rels)?$/.test(name)
    ),
    binaryParts: chart.binaryParts?.filter((part) => part.name.endsWith('Workbook1.xlsx')),
  };
}
