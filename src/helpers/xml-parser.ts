/**
 * Copyright (c) Microsoft Corporation.  All rights reserved.
 */

export function stringToXML(oString) {
    return (new DOMParser()).parseFromString(oString, 'text/xml');
}
