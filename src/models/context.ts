/**
 * Copyright (c) Microsoft Corporation.  All rights reserved.
 */

/**
 * Information about the query, included with the action context (IActionContext)
 */
export interface IQueryObject {
    id: string;
    isPublic: boolean;
    name: string;
    path: string;
    wiql: string;
}

/**
 * Incoming context sent when VSS calls getMenuItems (in app.ts)
 */
export interface IActionContext {
    id?: number;
    // From work item form
    workItemId?: number;
    query?: IQueryObject;
    queryText?: string;
    ids?: number[];
    // From backlog/iteration (context menu) and query results (toolbar and context menu)
    workItemIds?: number[];
    columns?: string[];
}
