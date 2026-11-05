/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Office.js `ThemeColorScheme`: the `a:clrScheme` of the theme that the master of a slide, a
// layout, or a master relates to (CT_ColorScheme, ECMA-376 Part 1, 20.1.6.2).

import { ClientResult } from '../runtime/client-result.ts';
import { hydrated } from '../runtime/hydrate.ts';
import type { ObjectPath } from '../runtime/object-path.ts';
import type { RequestContext } from '../runtime/request-context.ts';
import { THEME_COLOR, type ThemeColor } from './enums.ts';
import { ModelObject, text } from './model-object.ts';

/**
 * The theme colors of a slide, a layout, or a master. Every slide, layout, and master that uses
 * the same theme shares them, so a write changes all of them.
 *
 * @public
 */
export class ThemeColorScheme extends ModelObject {
  readonly #owner: ObjectPath;

  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, owner: ObjectPath) {
    super(context, path);
    this.#owner = owner;
  }

  /** One theme color as `#RRGGBB`. `Mixed` and `None` are `InvalidArgument`. */
  getThemeColor(color: ThemeColor): ClientResult<string>;
  /** One theme color as `#RRGGBB`. `Mixed` and `None` are `InvalidArgument`. */
  getThemeColor(
    color:
      | 'Accent1'
      | 'Accent2'
      | 'Accent3'
      | 'Accent4'
      | 'Accent5'
      | 'Accent6'
      | 'Dark1'
      | 'Dark2'
      | 'FollowedHyperlink'
      | 'Hyperlink'
      | 'Light1'
      | 'Light2'
      | 'Mixed'
      | 'None'
  ): ClientResult<string>;
  getThemeColor(color: ThemeColor | `${ThemeColor}`): ClientResult<string> {
    const label = this.at('getThemeColor');
    const slot = THEME_COLOR.toFile(color, label);
    const { result, fill } = ClientResult.create<string>(label);
    const owner = this.#owner;
    this.enqueue(
      'read',
      'getThemeColor',
      (ref) => ({ op: 'readThemeColor', owner: ref(owner), slot }),
      (value) => fill(hydrated(value, 'text', label).text)
    );
    return result;
  }

  /** Set one theme color: `#RRGGBB` or a preset color name. */
  setThemeColor(color: ThemeColor, rgbColor: string): void;
  /** Set one theme color: `#RRGGBB` or a preset color name. */
  setThemeColor(
    color:
      | 'Accent1'
      | 'Accent2'
      | 'Accent3'
      | 'Accent4'
      | 'Accent5'
      | 'Accent6'
      | 'Dark1'
      | 'Dark2'
      | 'FollowedHyperlink'
      | 'Hyperlink'
      | 'Light1'
      | 'Light2'
      | 'Mixed'
      | 'None',
    rgbColor: string
  ): void;
  setThemeColor(color: ThemeColor | `${ThemeColor}`, rgbColor: string): void {
    const slot = THEME_COLOR.toFile(color, this.at('setThemeColor'));
    const value = text(rgbColor, this.at('rgbColor'));
    const owner = this.#owner;
    this.command('setThemeColor', (ref) => ({
      op: 'setThemeColor',
      owner: ref(owner),
      slot,
      color: value,
    }));
  }
}
