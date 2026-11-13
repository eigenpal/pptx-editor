// SVG path data for evaluated geometry, and path commands mapped through a matrix: the slide
// paint, the shape gallery, and the soft edge of a group (`group-text-covers.ts`) read them.

import {
  apply as applyMatrix,
  type GeometryPath,
  type Matrix,
  type PathCommand,
} from '../layout/index.ts';
import { num } from './svg.ts';

/** Path data for an evaluated path in local EMU. */
export function pathData(path: GeometryPath): string {
  let d = '';
  for (const c of path.commands) {
    switch (c.op) {
      case 'M':
      case 'L':
        d += `${c.op}${num(c.x)} ${num(c.y)}`;
        break;
      case 'C':
        d += `C${num(c.x1)} ${num(c.y1)} ${num(c.x2)} ${num(c.y2)} ${num(c.x)} ${num(c.y)}`;
        break;
      case 'Q':
        d += `Q${num(c.x1)} ${num(c.y1)} ${num(c.x)} ${num(c.y)}`;
        break;
      case 'Z':
        d += 'Z';
        break;
    }
  }
  return d;
}

/** Path commands mapped through a matrix. */
export function mapCommands(commands: readonly PathCommand[], m: Matrix): PathCommand[] {
  return commands.map((c): PathCommand => {
    switch (c.op) {
      case 'Z':
        return c;
      case 'M':
      case 'L': {
        const p = applyMatrix(m, c.x, c.y);
        return { op: c.op, x: p.x, y: p.y };
      }
      case 'Q': {
        const [p1, p] = [applyMatrix(m, c.x1, c.y1), applyMatrix(m, c.x, c.y)];
        return { op: 'Q', x1: p1.x, y1: p1.y, x: p.x, y: p.y };
      }
      case 'C': {
        const [p1, p2] = [applyMatrix(m, c.x1, c.y1), applyMatrix(m, c.x2, c.y2)];
        const p = applyMatrix(m, c.x, c.y);
        return { op: 'C', x1: p1.x, y1: p1.y, x2: p2.x, y2: p2.y, x: p.x, y: p.y };
      }
    }
  });
}

/** SVG path data for path commands. */
export function commandData(commands: readonly PathCommand[]): string {
  return pathData({ commands, fill: 'none', stroke: true });
}
