/**
 * Copyright (c) Microsoft Corporation.  All rights reserved.
 */

import { QueryExpand, QueryHierarchyItem } from 'TFS/WorkItemTracking/Contracts';
import * as WorkItemTrackingClient from 'TFS/WorkItemTracking/RestClient';

import { NotesService, parseQueryJson, stringToXML } from '../helpers';
import { IActionContext, ODataMetadataParser } from '../models';
import { ODataParser } from '../parsers';

export class DialogHelper {
    private actionContext: IActionContext;
    private webContext: WebContext;

    constructor(context: IActionContext) {
        this.actionContext = context;
        this.webContext = VSS.getWebContext();
    }

    /**
     * Main Entry point.
     * @returns The full url query string.
     */
    public async getODataUrl(): Promise<string> {
        let query: QueryHierarchyItem;
        let metadata: any;

        const queryErrorString = 'The query could not be found. Please make sure this is a valid query.';
        const axErrorString = 'Could not get ADO Analytics metadata for this account. Please make sure ADO Analytics is enabled.';

        NotesService.instance.clearAllNotes();

        // Validate
        try {
            query = await this.getQueryTemp();
        } catch (ex) {
            NotesService.instance.newNote('error', queryErrorString);
            return null;
        }

        try {
            metadata = await this.getODataMetadata();
        } catch (e) {
            NotesService.instance.newNote('error', axErrorString);
            return null;
        }

        if (!(query != null && query.wiql != null)) {
            NotesService.instance.newNote('error', queryErrorString);
            return null;
        }

        if (metadata == null) {
            NotesService.instance.newNote('error', axErrorString);
            return null;
        }

        const parser = new ODataParser(query, metadata);

        if (!parser.isQueryValid()) {
            return null;
        }

        // Try to parse and return query.
        try {
            const allStatements = parser.makeAllStatements();

            // Test length.
            if (encodeURI(allStatements).length > 2083) {
                NotesService.instance.newNote('warning', `The encoded url has ${encodeURI(allStatements).length} characters and will be too long for some browsers and HTTP clients.`);
            }

            return allStatements;
        } catch (e) {
            NotesService.instance.newNote('error', `An unexpected error occured: ${e.message}`);
            return null;
        }
    }

    private async getQuery(): Promise<QueryHierarchyItem> {
        const workItemTrackingClient = WorkItemTrackingClient.getClient();
        const query = await workItemTrackingClient.getQuery(
            this.webContext.project.name,
            this.actionContext.query.id,
            QueryExpand.All
        );

        return query;
    }

    private async getQueryTemp(): Promise<QueryHierarchyItem> {
        const accessToken = await VSS.getAccessToken();
        const url = `${this.webContext.account.uri}/${this.webContext.project.name}/_apis/wit/queries/${this.actionContext.query.id}?%24expand=3`;
        const response = await window.fetch(url, {
            method: 'GET',
            headers: {
                Authorization: `Bearer ${accessToken.token}`,
            },
        });
        const responseText = await response.text();
        const responseObj = parseQueryJson(responseText);
        return responseObj;
    }

    private async getODataMetadata(): Promise<any> {
        const accessToken = await VSS.getAccessToken();
        const url = `https://analytics.dev.azure.com/${this.webContext.account.name}/${this.webContext.project.name}/_odata/v1.0/$metadata`;
        const response = await window.fetch(url, {
            method: 'GET',
            headers: {
                Authorization: `Bearer ${accessToken.token}`,
            },
        });

        if (response.status > 200 || response.status >= 400) {
            return null;
        }

        const responseString = await response.text();
        const responseXml = stringToXML(await responseString);
        return ODataMetadataParser.createPropertyMap(ODataMetadataParser.parseDocument(responseXml));
    }
}
