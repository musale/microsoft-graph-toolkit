/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Context, Middleware } from '@microsoft/microsoft-graph-client';
import { getRequestHeader, setRequestHeader } from '@microsoft/microsoft-graph-client/lib/es/middleware/MiddlewareUtil';
import { MgtBaseComponent } from '../components/baseComponent';
import { PACKAGE_VERSION } from './version';

/**
 * Implements Middleware for the Graph sdk to inject
 * the toolkit version in the SdkVersion header
 *
 * @class SdkVersionMiddleware
 * @implements {Middleware}
 */
export class SdkVersionMiddleware implements Middleware {
  /**
   * @private
   * A member to hold next middleware in the middleware chain
   */
  private _nextMiddleware: Middleware;
  private _component?: MgtBaseComponent;

  constructor(component?: MgtBaseComponent) {
    this._component = component;
  }

  // tslint:disable-next-line: completed-docs
  public async execute(context: Context): Promise<void> {
    try {
      // Header parts must follow the format: 'name/version'
      const headerParts: string[] = [];

      // Component version
      if (this._component) {
        const componentTagName = this._component.tagName.toLowerCase();
        const componentVersion: string = `${componentTagName}/${PACKAGE_VERSION}`;
        headerParts.push(componentVersion);
      }

      // Package version
      const packageVersion: string = `mgt/${PACKAGE_VERSION}`;
      headerParts.push(packageVersion);

      // Existing SdkVersion header value
      headerParts.push(getRequestHeader(context.request, context.options, 'SdkVersion'));

      // Join the header parts together and update the SdkVersion request header value
      const sdkVersionHeaderValue = headerParts.join(', ');
      setRequestHeader(context.request, context.options, 'SdkVersion', sdkVersionHeaderValue);

      return await this._nextMiddleware.execute(context);
    } catch (error) {
      throw error;
    }
  }
  /**
   * Handles setting of next middleware
   *
   * @param {Middleware} next
   * @memberof SdkVersionMiddleware
   */
  public setNext(next: Middleware): void {
    this._nextMiddleware = next;
  }
}