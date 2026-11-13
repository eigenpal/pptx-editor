import type { ColorChoice } from './neutral.ts';

/** A format target captured from one editor document. @public */
export interface ShapeFormatTarget {
  readonly documentId: string;
  readonly part: string;
  readonly nodeId: string;
}

/** One gradient stop. Opacity and position are fractions from 0 to 1. @public */
export interface GradientStopDetail {
  readonly id: string;
  readonly position: number;
  /** Absent when the source color cannot be represented by a color choice. */
  readonly color?: ColorChoice;
  readonly opacity: number;
}

/** A gradient without reducing its source color transforms. @public */
export interface GradientDetail {
  readonly kind: 'linear' | 'radial' | 'other';
  readonly angle: number;
  readonly rotWithShape: boolean;
  readonly stops: readonly GradientStopDetail[];
}

/** A solid fill. Opacity is a fraction from 0 to 1; transparency is one minus opacity. @public */
export interface SolidFillDetail {
  /** Absent when the source color cannot be represented by a color choice. */
  readonly color?: ColorChoice;
  readonly opacity: number;
}

/** A shadow. Lengths are points, angles are degrees, and opacity is a fraction. @public */
export interface ShadowDetail {
  readonly kind: 'inner' | 'outer';
  readonly blur: number;
  readonly distance: number;
  readonly angle: number;
  readonly color?: ColorChoice;
  readonly opacity: number;
  readonly rotWithShape: boolean;
}

/** A reflection. Lengths are points; positions and opacity are fractions. @public */
export interface ReflectionDetail {
  readonly blur: number;
  readonly distance: number;
  readonly startOpacity: number;
  readonly endOpacity: number;
  readonly startPosition: number;
  readonly endPosition: number;
  readonly rotWithShape: boolean;
}

/** Detailed formatting that partial writes can edit. @public */
export interface ShapeFormatDetail {
  readonly target: ShapeFormatTarget;
  /** The solid fill, own or inherited; null for any other fill. */
  readonly solid: SolidFillDetail | null;
  readonly gradient: GradientDetail | null;
  readonly shadow: ShadowDetail | null;
  readonly reflection: ReflectionDetail | null;
  /** Why effect details are unavailable. Other supported properties remain readable. */
  readonly effectsReason?: string;
}

/** Detailed formatting, or the reason the element cannot be inspected. @public */
export type ShapeFormatRead =
  | { readonly ok: true; readonly format: ShapeFormatDetail }
  | { readonly ok: false; readonly reason: string };

/** Change only specified properties. Other source XML remains unchanged. @public */
export type ShapeFormatPatch =
  | {
      readonly kind: 'solid';
      /** The fill opacity. The color and its other transforms stay as they are. */
      readonly opacity: number;
    }
  | {
      readonly kind: 'gradient';
      readonly angle?: number;
      readonly rotWithShape?: boolean;
      /** Add one stop. Existing stops keep their identities and color transforms. */
      readonly addStop?: {
        readonly position: number;
        readonly color: ColorChoice;
        readonly opacity?: number;
      };
      /** Remove one source stop. At least two stops must remain. */
      readonly removeStop?: string;
      readonly stop?: {
        readonly id: string;
        readonly position?: number;
        readonly color?: ColorChoice;
        readonly opacity?: number;
      };
    }
  | ({ readonly kind: 'shadow' } & Partial<Omit<ShadowDetail, 'kind'>>)
  | ({ readonly kind: 'reflection' } & Partial<ReflectionDetail>);

/** A partial formatting edit bound to the document and part that supplied its values. @public */
export interface PatchShapeFormatCommand {
  readonly type: 'patchShapeFormat';
  readonly target: ShapeFormatTarget;
  readonly patch: ShapeFormatPatch;
}

/** Detailed values without a browser document target. @public */
export type ShapeFormatProperties = Omit<ShapeFormatDetail, 'target'>;
