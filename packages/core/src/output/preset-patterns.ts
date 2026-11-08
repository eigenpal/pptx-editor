// The 54 DrawingML preset patterns (`ST_PresetPatternVal`, ECMA-376 Part 1, 20.1.10.51) as 8×8
// bitmaps (design D2 of `fill-and-line-fidelity`). Each pixel is the foreground's coverage: 1, 0.5,
// or 0. The specification gives no bitmaps. Every bitmap is read from the tiling pattern images of
// `fixtures/follow-ups.pdf`, slide 2, which fills one shape with each preset.

/**
 * One pattern pixel: 1 pt. The reference PDF draws each preset as an 8 × 8 pt tile
 * (`fixtures/follow-ups.pdf`, slide 2: pattern matrix scale 0.125 over a 64-unit cell).
 */
export const PATTERN_PIXEL = 12700;

/**
 * The bitmaps, in the schema's order, read from the reference PDF's 64 × 64 pattern images (8 × 8
 * image pixels per pattern pixel): `#` is the foreground, `+` the midpoint of the two colors (the
 * reference draws the diagonals of `dnDiag`, `upDiag`, and `diagCross` with half-tone neighbors),
 * and a space the background. Row 0 is the top row.
 */
// prettier-ignore
const BITMAPS: Readonly<Record<string, readonly string[]>> = {
  pct5: ['#       ', '        ', '        ', '        ', '    #   ', '        ', '        ', '        '],
  pct10: ['#       ', '        ', '    #   ', '        ', '#       ', '        ', '    #   ', '        '],
  pct20: ['#   #   ', '        ', '  #   # ', '        ', '#   #   ', '        ', '  #   # ', '        '],
  pct25: ['#   #   ', '  #   # ', '#   #   ', '  #   # ', '#   #   ', '  #   # ', '#   #   ', '  #   # '],
  pct30: ['# # # # ', ' #   #  ', '# # # # ', '   #   #', '# # # # ', ' #   #  ', '# # # # ', '   #   #'],
  pct40: ['# # # # ', ' # # # #', '# # # # ', ' # #   #', '# # # # ', ' # # # #', '# # # # ', '   # # #'],
  pct50: ['# # # # ', ' # # # #', '# # # # ', ' # # # #', '# # # # ', ' # # # #', '# # # # ', ' # # # #'],
  pct60: ['### ### ', ' # # # #', '# ### ##', ' # # # #', '### ### ', ' # # # #', '# ### ##', ' # # # #'],
  pct70: [' ### ###', '## ### #', ' ### ###', '## ### #', ' ### ###', '## ### #', ' ### ###', '## ### #'],
  pct75: [' ### ###', '########', '## ### #', '########', ' ### ###', '########', '## ### #', '########'],
  pct80: ['### ####', '########', '####### ', '########', '### ####', '########', '####### ', '########'],
  pct90: ['########', '########', '########', '#### ###', '########', '########', '########', ' #######'],
  horz: ['########', '        ', '        ', '        ', '        ', '        ', '        ', '        '],
  vert: ['#       ', '#       ', '#       ', '#       ', '#       ', '#       ', '#       ', '#       '],
  ltHorz: ['########', '        ', '        ', '        ', '########', '        ', '        ', '        '],
  ltVert: ['#   #   ', '#   #   ', '#   #   ', '#   #   ', '#   #   ', '#   #   ', '#   #   ', '#   #   '],
  dkHorz: ['########', '########', '        ', '        ', '########', '########', '        ', '        '],
  dkVert: ['##  ##  ', '##  ##  ', '##  ##  ', '##  ##  ', '##  ##  ', '##  ##  ', '##  ##  ', '##  ##  '],
  narHorz: ['########', '        ', '########', '        ', '########', '        ', '########', '        '],
  narVert: [' # # # #', ' # # # #', ' # # # #', ' # # # #', ' # # # #', ' # # # #', ' # # # #', ' # # # #'],
  dashHorz: ['####    ', '        ', '        ', '        ', '    ####', '        ', '        ', '        '],
  dashVert: ['#       ', '#       ', '#       ', '#       ', '    #   ', '    #   ', '    #   ', '    #   '],
  cross: ['########', '#       ', '#       ', '#       ', '#       ', '#       ', '#       ', '#       '],
  dnDiag: ['#+     +', '+#+     ', ' +#+    ', '  +#+   ', '   +#+  ', '    +#+ ', '     +#+', '+     +#'],
  upDiag: ['+     +#', '     +#+', '    +#+ ', '   +#+  ', '  +#+   ', ' +#+    ', '+#+     ', '#+     +'],
  ltDnDiag: ['#   #   ', ' #   #  ', '  #   # ', '   #   #', '#   #   ', ' #   #  ', '  #   # ', '   #   #'],
  ltUpDiag: ['   #   #', '  #   # ', ' #   #  ', '#   #   ', '   #   #', '  #   # ', ' #   #  ', '#   #   '],
  dkDnDiag: ['##  ##  ', ' ##  ## ', '  ##  ##', '#  ##  #', '##  ##  ', ' ##  ## ', '  ##  ##', '#  ##  #'],
  dkUpDiag: ['  ##  ##', ' ##  ## ', '##  ##  ', '#  ##  #', '  ##  ##', ' ##  ## ', '##  ##  ', '#  ##  #'],
  wdDnDiag: ['##     #', '###     ', ' ###    ', '  ###   ', '   ###  ', '    ### ', '     ###', '#     ##'],
  wdUpDiag: ['#     ##', '     ###', '    ### ', '   ###  ', '  ###   ', ' ###    ', '###     ', '##     #'],
  dashDnDiag: ['        ', '        ', '#   #   ', ' #   #  ', '  #   # ', '   #   #', '        ', '        '],
  dashUpDiag: ['        ', '        ', '   #   #', '  #   # ', ' #   #  ', '#   #   ', '        ', '        '],
  diagCross: ['#+    +#', '+#+  +#+', ' +#++#+ ', '  +##+  ', '  +##+  ', ' +#++#+ ', '+#+  +#+', '#+    +#'],
  smCheck: ['#  ##  #', ' ##  ## ', ' ##  ## ', '#  ##  #', '#  ##  #', ' ##  ## ', ' ##  ## ', '#  ##  #'],
  lgCheck: ['####    ', '####    ', '####    ', '####    ', '    ####', '    ####', '    ####', '    ####'],
  smGrid: ['########', '#   #   ', '#   #   ', '#   #   ', '########', '#   #   ', '#   #   ', '#   #   '],
  lgGrid: ['########', '#       ', '#       ', '#       ', '#       ', '#       ', '#       ', '#       '],
  dotGrid: ['# # # # ', '        ', '#       ', '        ', '#       ', '        ', '#       ', '        '],
  smConfetti: ['#       ', '    #   ', ' #      ', '      # ', '   #    ', '       #', '  #     ', '     #  '],
  lgConfetti: ['# ##   #', '  ##    ', '      ##', '   ## ##', '## ##   ', '##      ', '    ##  ', '#   ## #'],
  horzBrick: ['########', '#       ', '#       ', '#       ', '########', '    #   ', '    #   ', '    #   '],
  diagBrick: ['       #', '      # ', '     #  ', '    #   ', '   ##   ', '  #  #  ', ' #    # ', '#      #'],
  solidDmnd: ['   #    ', '  ###   ', ' #####  ', '####### ', ' #####  ', '  ###   ', '   #    ', '        '],
  openDmnd: ['#     # ', ' #   #  ', '  # #   ', '   #    ', '  # #   ', ' #   #  ', '#     # ', '       #'],
  dotDmnd: ['#       ', '        ', '  #   # ', '        ', '    #   ', '        ', '  #   # ', '        '],
  plaid: ['# # # # ', ' # # # #', '# # # # ', ' # # # #', '####    ', '####    ', '####    ', '####    '],
  sphere: [' ### ###', '#   #  #', '#   ####', '#   ####', ' ### ###', '#  ##   ', '#####   ', '#####   '],
  weave: ['#   #   ', ' # # #  ', '  #   # ', ' #   # #', '#   #   ', '   # #  ', '  #   # ', ' # #   #'],
  divot: ['        ', '   #    ', '    #   ', '   #    ', '        ', '#       ', '       #', '#       '],
  shingle: ['      ##', '#    #  ', ' #  #   ', '  ##    ', '    ##  ', '      # ', '       #', '       #'],
  wave: ['        ', '   ##   ', '  #  # #', '##      ', '        ', '   ##   ', '  #  # #', '##      '],
  trellis: ['########', ' ##  ## ', '########', '#  ##  #', '########', ' ##  ## ', '########', '#  ##  #'],
  zigZag: ['#      #', ' #    # ', '  #  #  ', '   ##   ', '#      #', ' #    # ', '  #  #  ', '   ##   '],
};

/** Every preset pattern name. */
export const PRESET_PATTERNS: readonly string[] = Object.keys(BITMAPS);

/**
 * The 8×8 bitmap of a preset (rows of foreground coverage: 1, 0.5, or 0), or undefined for an
 * unknown name.
 */
export function patternBitmap(preset: string): number[][] | undefined {
  if (!Object.prototype.hasOwnProperty.call(BITMAPS, preset)) return undefined;
  return BITMAPS[preset].map((row) =>
    Array.from(row, (c) => (c === '#' ? 1 : c === '+' ? 0.5 : 0))
  );
}

/** The ink density: the mean foreground coverage. */
export function patternDensity(preset: string): number | undefined {
  const bits = patternBitmap(preset);
  return bits ? bits.flat().reduce((sum, v) => sum + v, 0) / 64 : undefined;
}
