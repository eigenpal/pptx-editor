// The zoom range of the slide canvas: one range for the zoom buttons, the zoom menu, the
// keyboard, Ctrl/Cmd + wheel, and `setZoom`, so no input reaches a level another cannot leave.

/** The smallest zoom, 10 %. @internal */
export const ZOOM_MIN = 0.1;

/** The largest zoom, 400 %. @internal */
export const ZOOM_MAX = 4;

/** A zoom kept in the range. @internal */
export function clampZoom(zoom: number): number {
  return Math.max(ZOOM_MIN, Math.min(ZOOM_MAX, zoom));
}

/** The zoom steps the zoom buttons of the status bar move through, from end to end. */
const ZOOM_STEPS: readonly number[] = [
  ZOOM_MIN,
  0.25,
  0.33,
  0.5,
  0.67,
  0.75,
  0.8,
  0.9,
  1,
  1.1,
  1.25,
  1.5,
  1.75,
  2,
  2.5,
  3,
  ZOOM_MAX,
];

/** The zoom one step in from `zoom` (`1`) or out (`-1`), or null at the end of the range. @internal */
export function zoomStep(zoom: number, direction: 1 | -1): number | null {
  const next =
    direction > 0
      ? ZOOM_STEPS.find((z) => z > zoom + 0.001)
      : [...ZOOM_STEPS].reverse().find((z) => z < zoom - 0.001);
  return next ?? null;
}

/**
 * Why a zoom button is off at `zoom`: the i18n key of the reason at the end of the range, or
 * null while it can step. @internal
 */
export function zoomLimitReason(
  zoom: number,
  direction: 1 | -1
): 'disabled.zoomMax' | 'disabled.zoomMin' | null {
  if (zoomStep(zoom, direction) !== null) return null;
  return direction > 0 ? 'disabled.zoomMax' : 'disabled.zoomMin';
}

/** The pixels one wheel line and one wheel page stand for (`WheelEvent.deltaMode` 1 and 2). */
const LINE_PX = 33;
const PAGE_PX = 100;

/** The largest scroll of one wheel event that counts: one mouse notch (`deltaY` 100). */
const NOTCH_PX = 100;

/**
 * The zoom factor of one Ctrl/Cmd + wheel event: 10 % for a mouse notch (`deltaY` 100, or 3
 * lines), and in proportion for the small deltas of a trackpad pinch, so the zoom is smooth
 * there. A larger delta counts as one notch. @internal
 */
export function wheelZoomFactor(deltaY: number, deltaMode: number): number {
  const px = deltaMode === 1 ? deltaY * LINE_PX : deltaMode === 2 ? deltaY * PAGE_PX : deltaY;
  const notch = Math.max(-NOTCH_PX, Math.min(NOTCH_PX, px));
  return Math.pow(1.1, -notch / NOTCH_PX);
}

/**
 * The zoom keys: Ctrl/Cmd + = (or +) zooms in by 25 %, Ctrl/Cmd + − (or _) zooms out by 20 %,
 * and Ctrl/Cmd + 0 fits the slide to the window. The factor, `fit`, or null for another key.
 * @internal
 */
export function zoomKeyAction(key: string, mod: boolean): number | 'fit' | null {
  if (!mod) return null;
  if (key === '=' || key === '+') return 1.25;
  if (key === '-' || key === '_') return 0.8;
  return key === '0' ? 'fit' : null;
}
