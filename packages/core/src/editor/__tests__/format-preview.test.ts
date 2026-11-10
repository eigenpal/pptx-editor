import { expect, test } from 'bun:test';
import { createGradientStopIdentity, gradientPreviewCss } from '../format-detail-fields.ts';

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

test('gradient identity retains survivors when one patch removes and adds a stop', () => {
  const identity = createGradientStopIdentity();
  const before = [0, 0.5, 1].map((position, index) => ({
    id: `source-${index}`,
    position,
    opacity: 1,
  }));
  const selected = identity.identity(before, 'source-2');
  const after = [0.5, 0.75, 1].map((position, index) => ({
    id: `local-${index}`,
    position,
    opacity: 1,
  }));
  identity.patched(before, after, {
    kind: 'gradient',
    removeStop: 'source-0',
    addStop: { position: 0.75, color: '#FF0000' },
  });
  expect(identity.identity(after, 'local-2')).toBe(selected);
  expect(identity.identity(after, 'local-1')).not.toBe(selected);
  expect(identity.identity(before, 'source-2')).toBe(selected);
  const moved = [0.75, 0.9, 1].map((position, index) => ({
    id: `moved-${index}`,
    position,
    opacity: 1,
  }));
  const middle = identity.identity(before, 'source-1');
  identity.patched(before, moved, {
    kind: 'gradient',
    removeStop: 'source-0',
    stop: { id: 'source-1', position: 0.9 },
    addStop: { position: 0.75, color: '#FF0000' },
  });
  expect(identity.identity(moved, 'moved-1')).toBe(middle);
  expect(identity.identity(moved, 'moved-2')).toBe(selected);
});
