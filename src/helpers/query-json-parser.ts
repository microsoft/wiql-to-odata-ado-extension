/**
 * Copyright (c) Microsoft Corporation.  All rights reserved.
 */

import { LinkQueryMode, LogicalOperation, QueryHierarchyItem, QueryType } from 'TFS/WorkItemTracking/Contracts';

export function parseQueryJson(queryJson: string): QueryHierarchyItem {
    return JSON.parse(queryJson, (key: string, value: string) => {
        let newValue: any = value;
        if (key.toLocaleLowerCase().includes('date')) {
            newValue = new Date(value);
        } else if (key === 'logicalOperator') {
            newValue = LogicalOperation[value.toLocaleUpperCase()];
        } else {
            switch (value) {
                case 'flat':
                    newValue = QueryType.Flat;
                    break;
                case 'oneHop':
                    newValue = QueryType.OneHop;
                    break;
                case 'tree':
                    newValue = QueryType.Tree;
                    break;
                case 'linksOneHopMustContain':
                    newValue = LinkQueryMode.LinksOneHopMustContain;
                    break;
                case 'linksOneHopMayContain':
                    newValue = LinkQueryMode.LinksOneHopMayContain;
                    break;
                case 'linksOneHopDoesNotContain':
                    newValue = LinkQueryMode.LinksOneHopDoesNotContain;
                    break;

                // We dont really care about any other value, since it's not supported.
                default:
                    newValue = value;
                    break;
            }
        }

        return newValue;
    });
}
