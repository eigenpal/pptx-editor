import { expect, test } from 'bun:test';
import { gradientPreviewCss } from '../format-detail-fields.ts';

test('gradient preview sorts crossed positions without changing source order', () => {
  const stops = [
    { id: 'first', position: 0.8, color: '#FF0000', opacity: 0.5 },
    { id: 'second', position: 0.2, color: '#0000FF', opacity: 1 },
  ];
  expect(gradientPreviewCss(stops)).toBe(
    'linear-gradient(90deg, rgba(0,0,255,1) 20%, rgba(255,0,0,0.5) 80%)'
  );
  expect(stops.map((s) => s.id)).toEqual(['first', 'second']);
});

test('gradient preview does not invent colors for unsupported source values', () => {
  expect(
    gradientPreviewCss([{ id: 'unknown', position: 0, color: null, opacity: 1 }])
  ).toBeUndefined();
});
