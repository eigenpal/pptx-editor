// The 54 DrawingML preset patterns (`ST_PresetPatternVal`, ECMA-376 Part 1, 20.1.10.51) as 8×8
// bitmaps (design D2 of `fill-and-line-fidelity`). Each pixel is the foreground's coverage: 1, 0.5,
// or 0. The specification gives no bitmaps. Twelve presets take the bitmaps of the tiling patterns
// in `fixtures/fills-lines.pdf` (`MEASURED`). The others are an APPROXIMATION (listed on the public
// fidelity page): the percentage patterns come from an ordered-dither (Bayer) matrix, so their ink
// density is exact, and the rest are drawn to their names.

/**
 * One pattern pixel: 1 pt. The reference PDF draws each preset as an 8 × 8 pt tile
 * (`fixtures/fills-lines.pdf`, slide 1: pattern matrix scale 0.125 over a 64-unit cell).
 */
export const PATTERN_PIXEL = 12700;

/**
 * The bitmaps of the presets on `fixtures/fills-lines.pdf`, slide 1, read from its tiling pattern
 * images: `#` is the foreground, `+` the midpoint of the two colors (the reference draws the
 * diagonals with half-tone neighbors), and a space the background.
 */
// prettier-ignore
const MEASURED:Readonly<Record<string, readonly string[]>> = {
  pct25: ['#   #   ', '  #   # ', '#   #   ', '  #   # ', '#   #   ', '  #   # ', '#   #   ', '  #   # '],
  pct75: [' ### ###', '########', '## ### #', '########', ' ### ###', '########', '## ### #', '########'],
  dnDiag: ['#+     +', '+#+     ', ' +#+    ', '  +#+   ', '   +#+  ', '    +#+ ', '     +#+', '+     +#'],
  upDiag: ['+     +#', '     +#+', '    +#+ ', '   +#+  ', '  +#+   ', ' +#+    ', '+#+     ', '#+     +'],
  cross: ['########', '#       ', '#       ', '#       ', '#       ', '#       ', '#       ', '#       '],
  diagCross: ['#+    +#', '+#+  +#+', ' +#++#+ ', '  +##+  ', '  +##+  ', ' +#++#+ ', '+#+  +#+', '#+    +#'],
  smCheck: ['#  ##  #', ' ##  ## ', ' ##  ## ', '#  ##  #', '#  ##  #', ' ##  ## ', ' ##  ## ', '#  ##  #'],
  lgCheck: ['####    ', '####    ', '####    ', '####    ', '    ####', '    ####', '    ####', '    ####'],
  horzBrick: ['########', '#       ', '#       ', '#       ', '########', '    #   ', '    #   ', '    #   '],
  wave: ['        ', '   ##   ', '  #  # #', '##      ', '        ', '   ##   ', '  #  # #', '##      '],
  dkHorz: ['########', '########', '        ', '        ', '########', '########', '        ', '        '],
  dotGrid: ['# # # # ', '        ', '#       ', '        ', '#       ', '        ', '#       ', '        '],
};

/** Whether a preset's bitmap is measured on a reference PDF rather than approximated. */
export function isMeasuredPattern(preset: string): boolean {
  return Object.prototype.hasOwnProperty.call(MEASURED, preset);
}

const BAYER: readonly (readonly number[])[] = [
  [0, 32, 8, 40, 2, 34, 10, 42],
  [48, 16, 56, 24, 50, 18, 58, 26],
  [12, 44, 4, 36, 14, 46, 6, 38],
  [60, 28, 52, 20, 62, 30, 54, 22],
  [3, 35, 11, 43, 1, 33, 9, 41],
  [51, 19, 59, 27, 49, 17, 57, 25],
  [15, 47, 7, 39, 13, 45, 5, 37],
  [63, 31, 55, 23, 61, 29, 53, 21],
];

type Rule = (x: number, y: number) => boolean;

const percent =
  (p: number): Rule =>
  (x, y) =>
    BAYER[y][x] < Math.round((p / 100) * 64);
const dn = (x: number, y: number, n: number): number => (((x - y) % n) + n) % n;
const up = (x: number, y: number, n: number): number => (x + y) % n;
const points =
  (list: readonly [number, number][]): Rule =>
  (x, y) =>
    list.some(([px, py]) => px === x && py === y);
const WAVE = [2, 1, 1, 2, 3, 4, 4, 3];
const ZIG = [0, 1, 2, 3, 3, 2, 1, 0];

const RULES: Readonly<Record<string, Rule>> = {
  pct5: percent(5),
  pct10: percent(10),
  pct20: percent(20),
  pct25: percent(25),
  pct30: percent(30),
  pct40: percent(40),
  pct50: percent(50),
  pct60: percent(60),
  pct70: percent(70),
  pct75: percent(75),
  pct80: percent(80),
  pct90: percent(90),
  horz: (_x, y) => y === 0,
  vert: (x) => x === 0,
  ltHorz: (_x, y) => y % 4 === 0,
  ltVert: (x) => x % 4 === 0,
  dkHorz: (_x, y) => y % 4 < 2,
  dkVert: (x) => x % 4 < 2,
  narHorz: (_x, y) => y % 2 === 0,
  narVert: (x) => x % 2 === 0,
  dashHorz: (x, y) => (y === 0 && x < 4) || (y === 4 && x >= 4),
  dashVert: (x, y) => (x === 0 && y < 4) || (x === 4 && y >= 4),
  cross: (x, y) => x === 0 || y === 0,
  dnDiag: (x, y) => dn(x, y, 8) === 0,
  upDiag: (x, y) => up(x, y, 8) === 7,
  ltDnDiag: (x, y) => dn(x, y, 4) === 0,
  ltUpDiag: (x, y) => up(x, y, 4) === 3,
  dkDnDiag: (x, y) => dn(x, y, 4) < 2,
  dkUpDiag: (x, y) => up(x, y, 4) >= 2,
  wdDnDiag: (x, y) => dn(x, y, 8) < 3,
  wdUpDiag: (x, y) => up(x, y, 8) >= 5,
  dashDnDiag: (x, y) => dn(x, y, 4) === 0 && y % 8 < 4,
  dashUpDiag: (x, y) => up(x, y, 4) === 3 && y % 8 < 4,
  diagCross: (x, y) => dn(x, y, 8) === 0 || up(x, y, 8) === 7,
  smCheck: (x, y) => ((x >> 1) + (y >> 1)) % 2 === 0,
  lgCheck: (x, y) => ((x >> 2) + (y >> 2)) % 2 === 0,
  smGrid: (x, y) => x % 4 === 0 || y % 4 === 0,
  lgGrid: (x, y) => x === 0 || y === 0,
  dotGrid: (x, y) => (y === 0 && x % 2 === 0) || (x === 0 && y % 2 === 0),
  smConfetti: points([
    [0, 0],
    [4, 1],
    [1, 3],
    [6, 4],
    [3, 5],
    [7, 7],
  ]),
  lgConfetti: (x, y) =>
    [
      [0, 0],
      [4, 2],
      [1, 5],
      [6, 6],
    ].some(([px, py]) => (x - px + 8) % 8 < 2 && (y - py + 8) % 8 < 2),
  horzBrick: (x, y) => y === 0 || y === 4 || (x === 0 && y < 4) || (x === 4 && y > 4),
  diagBrick: (x, y) => up(x, y, 8) === 7 || (dn(x, y, 8) === 0 && x < 4),
  solidDmnd: (x, y) => Math.abs(x - 3.5) + Math.abs(y - 3.5) <= 3.5,
  openDmnd: (x, y) => dn(x, y, 8) === 0 || up(x, y, 8) === 7,
  dotDmnd: (x, y) => (dn(x, y, 8) === 0 || up(x, y, 8) === 0) && x % 2 === 0,
  plaid: (x, y) => (y < 4 && (x + y) % 2 === 0) || (x < 4 && y >= 4),
  sphere: (x, y) => (x - 3.5) ** 2 + (y - 3.5) ** 2 >= 12,
  weave: (x, y) => (up(x, y, 4) === 0 && y < 4) || (dn(x, y, 4) === 0 && y >= 4),
  divot: points([
    [2, 1],
    [3, 2],
    [2, 3],
    [6, 5],
    [7, 6],
    [6, 7],
  ]),
  shingle: (x, y) => (up(x, y, 8) === 7 && x < 6) || (y === 1 && x >= 6),
  wave: (x, y) => y === WAVE[x],
  trellis: (x, y) => (y % 2 === 0 ? (x + (y % 4)) % 4 !== 3 : (x + (y % 4)) % 4 === 1),
  zigZag: (x, y) => y % 4 === ZIG[x] % 4,
};

/** Every preset pattern name. */
export const PRESET_PATTERNS: readonly string[] = Object.keys(RULES);

/**
 * The 8×8 bitmap of a preset (rows of foreground coverage: 1, 0.5, or 0), or undefined for an
 * unknown name.
 */
export function patternBitmap(preset: string): number[][] | undefined {
  if (isMeasuredPattern(preset))
    return MEASURED[preset].map((row) =>
      Array.from(row, (c) => (c === '#' ? 1 : c === '+' ? 0.5 : 0))
    );
  const rule = Object.prototype.hasOwnProperty.call(RULES, preset) ? RULES[preset] : undefined;
  if (!rule) return undefined;
  return Array.from({ length: 8 }, (_, y) =>
    Array.from({ length: 8 }, (_, x) => (rule(x, y) ? 1 : 0))
  );
}

/** The ink density: the mean foreground coverage. */
export function patternDensity(preset: string): number | undefined {
  const bits = patternBitmap(preset);
  return bits ? bits.flat().reduce((sum, v) => sum + v, 0) / 64 : undefined;
}
