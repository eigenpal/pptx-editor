/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Office.js `Application`: the object that stands for the host application. Office.js gives it
// no properties and no methods, so it has none here either; it exists so that
// `context.application` and `load()` on it work as in Office.js.

import type { ResolvedLoadOptions } from '../runtime/load-options.ts';
import { ObjectPath } from '../runtime/object-path.ts';
import type { RequestContext } from '../runtime/request-context.ts';
import { ModelObject } from './model-object.ts';

/**
 * The application the presentation is open in. It has no properties or methods.
 *
 * @public
 */
export class Application extends ModelObject {
  /** @internal */
  private constructor(context: RequestContext, path: ObjectPath) {
    super(context, path);
  }

  /** @internal The application of a request context. It addresses nothing in the host. */
  static open(context: RequestContext): Application {
    return new Application(context, ObjectPath.pending('application'));
  }

  /** @internal There is nothing to load. */
  protected override onLoad(request: ResolvedLoadOptions): void {
    this.selection(request, []);
  }
}
