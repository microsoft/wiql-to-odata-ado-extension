/**
 * Copyright (c) Microsoft Corporation.  All rights reserved.
 */

import { WorkItemFieldReference, WorkItemQueryClause } from 'TFS/WorkItemTracking/Contracts';

import { NotesService } from '../helpers';
import { IProperty, ODataMetadataParser } from '../models';

interface ISupportedOperationsHandling {
    name: string;
    handler: (clause: WorkItemQueryClause) => string;
}

interface ISupportedTypesHandling {
    name: string;
    mathAllowed: boolean;
    handler: (value: string) => string;
}

interface ILhsRhs {
    left: string;
    right: string;
}

/**
 * Handles the parsing of a single clause into an OData clause.
 */
export class ClauseParser {
    /**
     * A cache for the storage of the supportedOperationsHandling variable in a map with key = WIQL SupportedOperation name.
     * Managed by the supportedOperationsHandlingMap function.
     */
    private supportedOperationsHandlingMapInternal: { [id: string]: ISupportedOperationsHandling; };

    /**
     * A cache for the storage of the supportedTypesHandling variable in a map with key = the OData type name.
     * Managed by the supportedTypesHandlingMap function.
     */
    private supportedTypesHandlingMapInternal: { [id: string]: ISupportedTypesHandling; };

    /**
     * A cache for the storage the list of mathAllowed properties from the supportedTypesHandling variable.
     * Managed by the _mathAllowedTypes function.
     */
    private mathAllowedTypesInternal: string[];

    /**
     * A mapping of Ref.ReferenceNames to IProperties from IMetadata.
     */
    private metadata: { [id: string]: IProperty; };

    /**
     * VSS Web Context.
     */
    private webContext = VSS.getWebContext();

    /**
     * The WIQL operations and the functions to handle the parsing of the clause containing the operation.
     */
    private supportedOperationsHandling: ISupportedOperationsHandling[] = [
        { name: 'SupportedOperations.Contains', handler: (clause) => this.parseContains(clause) },
        { name: 'SupportedOperations.ContainsWords', handler: (clause) => this.parseContainsWord(clause) },
        { name: 'SupportedOperations.NotContainsWords', handler: (clause) => this.parseContainsWord(clause) },
        { name: 'SupportedOperations.Equals', handler: (clause) => this.parseSimpleOperator(clause, 'eq') },
        { name: 'SupportedOperations.EqualsField', handler: (clause) => this.parseSimpleOperator(clause, 'eq') },
        { name: 'SupportedOperations.Ever', handler: (clause) => this.parseEver(clause) },
        { name: 'SupportedOperations.GreaterThan', handler: (clause) => this.parseSimpleOperator(clause, 'gt') },
        { name: 'SupportedOperations.GreaterThanEquals', handler: (clause) => this.parseSimpleOperator(clause, 'ge') },
        { name: 'SupportedOperations.GreaterThanEqualsField', handler: (clause) => this.parseSimpleOperator(clause, 'ge') },
        { name: 'SupportedOperations.GreaterThanField', handler: (clause) => this.parseSimpleOperator(clause, 'gt') },
        { name: 'SupportedOperations.In', handler: (clause) => this.parseIn(clause) },
        // { name: 'SupportedOperations.InGroup', handler: () => {} }, // Not Supported by OData.
        { name: 'SupportedOperations.LessThan', handler: (clause) => this.parseSimpleOperator(clause, 'lt') },
        { name: 'SupportedOperations.LessThanEquals', handler: (clause) => this.parseSimpleOperator(clause, 'le') },
        { name: 'SupportedOperations.LessThanEqualsField', handler: (clause) => this.parseSimpleOperator(clause, 'le') },
        { name: 'SupportedOperations.LessThanField', handler: (clause) => this.parseSimpleOperator(clause, 'lt') },
        { name: 'SupportedOperations.NotContains', handler: (clause) => `not ${this.parseContains(clause)}` },
        { name: 'SupportedOperations.NotEquals', handler: (clause) => this.parseSimpleOperator(clause, 'ne') },
        { name: 'SupportedOperations.NotEqualsField', handler: (clause) => this.parseSimpleOperator(clause, 'ne') },
        { name: 'SupportedOperations.NotIn', handler: (clause) => `not ${this.parseIn(clause)}` },
        // { name: 'SupportedOperations.NotInGroup', handler: () => {} }, // Not Supported by OData.
        { name: 'SupportedOperations.NotUnder', handler: (clause) => `not ${this.parseUnder(clause)}` },
        { name: 'SupportedOperations.Under', handler: (clause) => this.parseUnder(clause) },
    ];

    /**
     * The metadata for each OData type and the functions to validate and parse the WIQL values into OData values.
     */
    private supportedTypesHandling: ISupportedTypesHandling[] = [
        { name: 'Edm.Binary', mathAllowed: false, handler: (value) => this.parseWithQuotesAndPrefix(value, 'X') },
        { name: 'Edm.Boolean', mathAllowed: false, handler: (value) => this.validateAndThen(/^true|false$/, value, () => this.parseWithoutQuotes(value)) },
        { name: 'Edm.Byte', mathAllowed: true, handler: (value) => this.parseWithoutQuotes(value) },
        { name: 'Edm.DateTime', mathAllowed: false, handler: (value) => this.parseDateTime(value) },
        { name: 'Edm.Decimal', mathAllowed: true, handler: (value) => this.validateAndThen(/^-?[0-9]+.[0-9]+$/, value, () => this.parseWithoutQuotesAndSuffix(value, 'M')) },
        { name: 'Edm.Double', mathAllowed: true, handler: (value) => this.validateAndThen(/^-?[0-9]+.[0-9]+$/, value, () => this.parseWithoutQuotesAndSuffix(value, 'd')) },
        { name: 'Edm.Single', mathAllowed: true, handler: (value) => this.validateAndThen(/^-?[0-9]+.[0-9]+$/, value, () => this.parseWithoutQuotesAndSuffix(value, 'f')) },
        {
            name: 'Edm.Guid',
            mathAllowed: false,
            handler: (value) => this.validateAndThen(/^[{(]?[0-9A-F]{8}[-]?(?:[0-9A-F]{4}[-]?){3}[0-9A-F]{12}[)}]?$^-?[0-9]+.[0-9]$/, value, () => this.parseWithQuotesAndPrefix(value, 'guid')),
        },
        { name: 'Edm.Int16', mathAllowed: true, handler: (value) => this.validateAndThen(/^-?[0-9]+$/, value, () => this.parseWithoutQuotes(value)) },
        { name: 'Edm.Int32', mathAllowed: true, handler: (value) => this.validateAndThen(/^-?[0-9]+$/, value, () => this.parseWithoutQuotes(value)) },
        { name: 'Edm.Int64', mathAllowed: true, handler: (value) => this.validateAndThen(/^-?[0-9]+$/, value, () => this.parseWithoutQuotesAndSuffix(value, 'L')) },
        { name: 'Edm.String', mathAllowed: false, handler: (value) => this.parseWithQuotesAndPrefix(value, '') },
        { name: 'Edm.DateTimeOffset', mathAllowed: false, handler: (value) => this.parseDateTimeOffset(value) },
        { name: 'Microsoft.VisualStudio.Services.Analytics.Model.CalendarDate', mathAllowed: false, handler: (value) => this.parseDateTimeOffset(value) },
        { name: 'Microsoft.VisualStudio.Services.Analytics.Model.Project', mathAllowed: false, handler: (value) => this.parseWithQuotesAndPrefix(value, '') },
        { name: 'Microsoft.VisualStudio.Services.Analytics.Model.User', mathAllowed: false, handler: (value) => this.parseEmailUser(value) },
    ];

    public constructor(metadata: { [id: string]: IProperty; }) {
        this.metadata = metadata;
    }

    /**
     * The entry for the clause parser.
     * @param clause the individual clause to be parsed to an OData string
     * @returns The OData clause if it can be parsed, else null. In the case of null, notes are added to the NotesService as to why.
     */
    public parseClause(clause: WorkItemQueryClause): string {
        const handling = this.supportedOperationsHandlingMap[clause.operator.referenceName];

        if (handling == null) {
            // Verify operator.
            NotesService.instance.newNote('warning', `The operation '${clause.operator.name}' is not supported by OData. Leaving the clause out of $filter.`);
            return null;
        }

        // Verify fields. There is always a field on the left hand side, and sometimes on the right hand side.
        const lhsMetadata = this.metadata[clause.field.referenceName];
        const rhsMetadata = clause.isFieldValue ? this.metadata[clause.fieldValue.referenceName] : this.metadata[clause.field.referenceName];
        if (lhsMetadata == null || (clause.isFieldValue && rhsMetadata)) {
            NotesService.instance.newNote(
                'warning',
                `There is no OData equivalent property for field ${lhsMetadata == null ? clause.field.referenceName : clause.fieldValue.referenceName}. Leaving the clause out of $filter.`
            );
            return null;
        }

        try {
            return handling.handler(clause);
        } catch (e) {
            NotesService.instance.newNote('warning', `There was a problem parsing the clause for field ${clause.field.referenceName}: ${e.message}. Leaving the clause out of $filter.`);
            return null;
        }
    }

    /**
     * For the storage of the supportedOperationsHandling variable in a map with key = WIQL SupportedOperation name.
     */
    private get supportedOperationsHandlingMap(): { [id: string]: ISupportedOperationsHandling; } {
        if (this.supportedOperationsHandlingMapInternal == null) {
            this.supportedOperationsHandlingMapInternal = this.supportedOperationsHandling.reduce<{ [name: string]: ISupportedOperationsHandling; }>(
                (acc, x) => { acc[x.name] = x; return acc; },
                {}
            );
        }
        return this.supportedOperationsHandlingMapInternal;
    }

    /**
     * For the storage of the supportedTypesHandling variable in a map with key = the OData type name.
     */
    private get supportedTypesHandlingMap(): { [id: string]: ISupportedTypesHandling; } {
        if (this.supportedTypesHandlingMapInternal == null) {
            this.supportedTypesHandlingMapInternal = this.supportedTypesHandling.reduce<{ [name: string]: ISupportedTypesHandling; }>(
                (acc, x) => { acc[x.name] = x; return acc; },
                {}
            );
        }
        return this.supportedTypesHandlingMapInternal;
    }

    /**
     * For the storage the list of mathAllowed properties from the supportedTypesHandling variable.
     */
    private get mathAllowedTypes(): string[] {
        if (this.mathAllowedTypesInternal == null) {
            this.mathAllowedTypesInternal = this.supportedTypesHandling.filter((h) => h.mathAllowed === true).map((h) => h.name);
        }
        return this.mathAllowedTypesInternal;
    }

    /**
     * Gets the left hand side and the right hand side of the OData clause, converting it from WIQL field names and values.
     */
    private getLhsRhs(clause: WorkItemQueryClause): ILhsRhs {
        const lhs = this.odataExpressionFromField(clause.field);
        const rhs = clause.isFieldValue ? this.odataExpressionFromField(clause.fieldValue) : this.odataExpressionFromValue(clause.field, clause.value);
        return { right: rhs, left: lhs };
    }

    /**
     * @returns the string name of the OData property corresponding to the WIQL field.
     * @param field The field reference of the field to convert to an OData property.
     */
    private odataExpressionFromField(field: WorkItemFieldReference): string {
        const odataProperty = this.metadata[field.referenceName];

        if (odataProperty.type.startsWith('Edm')) {
            // This is a primative type. Just return the field name.
            return odataProperty.name;
        } else {
            // This is a complex type. Append default field.
            const defaultField = ODataMetadataParser.defaultFields[odataProperty.type];
            if (defaultField == null) {
                throw new Error(`There is no default field for the OData datatype for ${field.name}`);
            }
            return `${odataProperty.name}/${defaultField.defaultFieldName}`;
        }
    }

    /**
     * @returns the OData formatted value.
     * @param field The WIQL field of the value.
     * @param value The WIQL value.
     */
    private odataExpressionFromValue(field: WorkItemFieldReference, value: string): string {
        // first, clean up any special values.
        const macrosInValue = value.match(/^@([A-Za-z]+)/g) || [];
        macrosInValue.forEach((macro) => {
            switch (macro.toLocaleLowerCase()) {
                case '@me':
                    value = value.replace(macro, this.webContext.user.email);
                    NotesService.instance.newNote('warning', `There is no OData equivalent for @me. Replacing with the static value of '${value}'.`);
                    break;
                case '@project':
                    value = value.replace(macro, this.webContext.project.name);
                    NotesService.instance.newNote('warning', `There is no OData equivalent for @project. Replacing with the static value of '${value}'.`);
                    break;
                case '@today':
                    const dateText = value.trim().toLocaleLowerCase();
                    if (dateText.length === '@today'.length) {
                        // If there are no arithmetic operations, we can use now().
                        value = 'date(now())';
                    } else {
                        // Arithmetic operations are currently not supported by VSTS OData. Replace with static value.
                        const matches = dateText.match(/@today ?([-+]) ?([0-9]+)/);
                        const theDate = new Date(Date.now());
                        theDate.setDate(theDate.getDate() + parseInt(`${matches[1]}${matches[2]}`, 10));
                        value = theDate.toISOString();
                        NotesService.instance.newNote('warning', `Arithmetic operations on @today are not supported by OData at this time. Replacing with the static value of '${value}'.`);
                    }
                    break;
                default:
                    NotesService.instance.newNote('warning', `There is no OData equivalent for ${macro}. Please manually replace the value in your OData query.`);
                    break;
            }
        });

        const oDataType = this.metadata[field.referenceName].type;

        // Replace math operations for certain data types where math is allowed.
        if (this.mathAllowedTypes.includes(oDataType)) {
            value = value.replace(/ ?- ?/, ' sub ');
            value = value.replace(/ ?\+ ?/, ' add ');
            value = value.replace(/ ?\* ?/, ' mul ');
            value = value.replace(/ ?\/ ?/, ' div ');
        }

        // Send to handler for type.
        const handler = this.supportedTypesHandlingMap[oDataType];
        return handler.handler(value);
    }

    //#region parseOperations

    private parseSimpleOperator(clause: WorkItemQueryClause, operator: string): string {
        const rhslhs = this.getLhsRhs(clause);
        return `${rhslhs.left} ${operator} ${rhslhs.right}`;
    }

    private parseContains(clause: WorkItemQueryClause): string {
        const rhslhs = this.getLhsRhs(clause);
        return `contains(${rhslhs.left}, ${rhslhs.right})`;
    }

    private parseContainsWord(clause: WorkItemQueryClause): string {
        const rhslhs = this.getLhsRhs(clause);
        const rhs = rhslhs.right.substr(1).slice(0, -1); // remove quotes
        return `(contains(${rhslhs.left}, '${rhs} ') or contains(${rhslhs.left}, ' ${rhs}'))`;
    }

    private parseEver(clause: WorkItemQueryClause): string {
        const rhslhs = this.getLhsRhs(clause);
        return `Revisions/any(r:r/${rhslhs.left} eq ${rhslhs.right})`;
    }

    private parseUnder(clause: WorkItemQueryClause): string {
        const rhslhs = this.getLhsRhs(clause);
        return `startswith(${rhslhs.left}, ${rhslhs.right})`;
    }

    private parseIn(clause: WorkItemQueryClause): string {
        const lhs = this.odataExpressionFromField(clause.field);
        const values = clause.value
            .replace(/[\(\)]/, '')
            .split(',')
            .map((s) => s.trim());

        if (values.length === 0) {
            NotesService.instance.newNote('warning', `There was an issue parsing the in operator list for property ${clause.field.referenceName}. Leaving the clause out of $filter.`);
            return null;
        }

        const orStatements = values.map((s) => `${lhs} eq ${this.odataExpressionFromValue(clause.field, s)}`).join(' or ');

        return `(${orStatements})`;
    }

    //#endregion

    //#region parseTypes

    private validateAndThen(regexTest: RegExp, value: string, parser: () => string): string {
        if (!regexTest.test(value)) {
            throw new Error(`The value '${value}' is not formatted properly for the OData datatype`);
        }

        return parser();
    }

    private parseWithQuotesAndPrefix(value: string, prefix: string): string {
        return `${prefix}'${value}'`;
    }

    private parseWithoutQuotes(value: string): string {
        return value;
    }

    private parseDateTime(value: string): string {
        const date = new Date(value);
        return `datetime'${date.toISOString()}'`;
    }

    private parseDateTimeOffset(value: string): string {
        if (value.includes('date(now())')) {
            return value;
        } else {
            return new Date(value).toISOString();
        }
    }

    private parseWithoutQuotesAndSuffix(value: string, suffix: string): string {
        return `${value}${suffix}`;
    }

    private parseEmailUser(value: string): string {
        value = (value.match(/<([^<>]+)>/) || [])[1] || value;
        return `'${value}'`;
    }

    //#endregion
}
