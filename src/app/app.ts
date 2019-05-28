/**
 * Copyright (c) Microsoft Corporation.  All rights reserved.
 */

import { environment } from '@ms/configuration-environment';
import { IActionContext } from '../models';

export const openQueryAction = {
    getMenuItems: (context: IActionContext): IContributedMenuItem[] => {
        if (!context || !context.query || !context.query.wiql) {
            return null;
        }
        return [{
            title: environment.buttonText,
            text: environment.buttonText,
            icon: 'images/favicon16x16.png',
            action: (actionContext: IActionContext) => {
                if (actionContext != null && actionContext.query != null && actionContext.query.id != null) {
                    openDialog(actionContext);
                }
            },
        }];

    },
};

export const openQueryOnToolbarAction = {
    getMenuItems: (context: IActionContext): IContributedMenuItem[] => {
        return [{
            title: environment.buttonText,
            text: environment.buttonText,
            icon: 'images/favicon16x16.png',
            action: async (actionContext: IActionContext) => {
                if (actionContext && actionContext.query && actionContext.query.wiql) {
                    openDialog(actionContext);
                } else {
                    const hostDialogService = await VSS.getService<IHostDialogService>(VSS.ServiceIds.Dialog);
                    hostDialogService.openMessageDialog(
                        `In order to open your query, please save it first in "My Queries" or "Shared Queries".`,
                        {
                            title: 'Unable to perform this operation',
                            buttons: [hostDialogService.buttons.ok],
                        });
                }
            },
        }];
    },
};

async function openDialog(actionContext: IActionContext) {
    const hostDialogService = await VSS.getService<IHostDialogService>(VSS.ServiceIds.Dialog);
    const dialog = await hostDialogService.openDialog(
        `${extensionContext.publisherId}.${extensionContext.extensionId}.app-dialog`,
        {
            title: `Translate to OData`,
            width: 500,
            height: 600,
            modal: true,
            draggable: true,
            resizable: true,
            buttons: {
                ok: {
                    id: 'ok',
                    text: 'Dismiss',
                    click: () => {
                        dialog.close();
                    },
                    class: 'cta',
                },
            },
        },
        actionContext
    );
}

const extensionContext = VSS.getExtensionContext();
VSS.register(
    `${extensionContext.publisherId}.${extensionContext.extensionId}.work-item-query-menu`,
    openQueryAction
);
VSS.register(
    `${extensionContext.publisherId}.${extensionContext.extensionId}.work-item-query-results-toolbar-menu`,
    openQueryOnToolbarAction
);
