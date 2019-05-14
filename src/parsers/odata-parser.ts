/**
 * Copyright (c) Microsoft Corporation.  All rights reserved.
 */

import { LinkQueryMode, LogicalOperation, QueryHierarchyItem, QueryRecursionOption, QueryType, WorkItemQueryClause } from 'TFS/WorkItemTracking/Contracts';

import { NotesService } from '../helpers';
import { IProperty, ODataMetadataParser } from '../models';
import { ClauseParser } from './clause-parsers';

interface ISelectAndExpand {
    select: string;
    expand: string;
}

/**
 * Parse a WIQL query into an OData query.
 */
export class ODataParser {
    /**
     * The root QueryHeirarchyItem.
     */
    private query: QueryHierarchyItem;

    /**
     * A mapping of Ref.ReferenceNames to IProperties from IMetadata.
     */
    private metadata: { [id: string]: IProperty; };

    /**
     * An instance of the Clause Parser.
     */
    private clauseParser: ClauseParser;

    /**
     * VSS Web Context.
     */
    private webContext = VSS.getWebContext();

    /**
     *
     * @param query The root QueryHeirarchyItem.
     * @param metadata A mapping of Ref.ReferenceNames to IProperties from IMetadata.
     */
    public constructor(query: QueryHierarchyItem, metadata: { [id: string]: IProperty; }) {
        this.query = query;
        this.metadata = metadata;
        this.clauseParser = new ClauseParser(metadata);
    }

    /**
     * Makes notes and returns a boolean denoting whether the query can be parsed or not.
     */
    public isQueryValid(): boolean {
        if (this.query.queryType === QueryType.Tree) {
            NotesService.instance.newNote('error', `Only 'Flat list of work items' and 'Work items and direct links' queries are supported by OData.`);
            return false;
        }

        if (this.query.isInvalidSyntax) {
            NotesService.instance.newNote('error', `The WIQL query has invalid syntax. Fix the query before continuing.`);
            return false;
        }
        return true;
    }

    /**
     * @returns A full url OData query that matches the WIQL query.
     */
    public makeAllStatements(): string {
        if (this.query.queryType === QueryType.Flat) {
            const selectAndExpand = this.makeSelectAndExpandStatement();
            const statements = [
                selectAndExpand.select,
                selectAndExpand.expand,
                this.makeFilterStatement(this.query.clauses),
                this.makeOrderByStatement(),
            ].filter((column) => column != null).join('&');
            return `https://analytics.dev.azure.com/${this.webContext.account.name}/_odata/v1.0/WorkItems?${statements}`;
        } else if (this.query.queryType === QueryType.OneHop) {
            NotesService.instance.newNote('info', `Non-Flat queries are not recommended by Azure DevOps Analytics. Expect warning VS403508 and possibly long response times.`);
            const selectAndExpand = this.makeSelectAndExpandStatementForOneHop();
            const statements = [
                selectAndExpand.select,
                selectAndExpand.expand,
                this.makeFilterStatementForOneHop(),
                this.makeOrderByStatement(),
            ].filter((column) => column != null).join('&');
            return `https://analytics.dev.azure.com/${this.webContext.account.name}/_odata/v1.0/WorkItems?${statements}`;
        }
    }

    public makeFilterStatementForOneHop(): string {
        const filterStatement = this.makeFilterStatement(this.query.sourceClauses);
        let linksQueryAdditionalClause: string = null;
        if (this.query.filterOptions !== LinkQueryMode.LinksOneHopMayContain) {
            // if not 'may contain', filter the base query down to only the work items in the target/link clause.
            const additionalFilterStatements = [
                this.makeFilterStatementSimple(this.query.targetClauses, 'l/TargetWorkItem'),
                this.makeFilterStatementSimple(this.query.linkClauses, 'l'),
            ].filter((column) => column != null).join(' and ');
            linksQueryAdditionalClause = additionalFilterStatements.length === 0 ? null :
                `${this.query.filterOptions === LinkQueryMode.LinksOneHopDoesNotContain ? 'not ' : ''}Links/any(l: ${additionalFilterStatements})`;
        }

        return filterStatement != null ?
            `${filterStatement}${linksQueryAdditionalClause != null ? ` and ${linksQueryAdditionalClause}` : ``}` :
            linksQueryAdditionalClause != null ? `$filter=${linksQueryAdditionalClause}` : null;
    }

    public makeSelectAndExpandStatementForOneHop(): ISelectAndExpand {
        const selectAndExpand = this.makeSelectAndExpandStatement();
        if (this.query.filterOptions !== LinkQueryMode.LinksOneHopDoesNotContain) {
            // include child expands when query requires them.
            const childrenStatements = [selectAndExpand.select, selectAndExpand.expand].filter((column) => column != null).join('; ');

            const filterStatements = [
                this.makeFilterStatementSimple(this.query.targetClauses, 'TargetWorkItem'),
                this.makeFilterStatementSimple(this.query.linkClauses),
            ].filter((column) => column != null).join(' and ');

            const children = `Links($select=TargetWorkItem; ${filterStatements.length === 0 ? '' : `$filter=${filterStatements}; `}$expand=TargetWorkItem(${childrenStatements}))`;
            return {
                select: selectAndExpand.select,
                expand: selectAndExpand.expand == null ? `$expand=${children}` : `${selectAndExpand.expand}, ${children}`,
            };
        }

        return selectAndExpand;
    }

    /**
     * @returns The $select statement, or null if it can't be parsed.
     */
    public makeSelectAndExpandStatement(): ISelectAndExpand {
        if (this.query.columns.length > 0) {
            const expandList: string[] = [];
            const selectList: string[] = [];

            this.query.columns.forEach((column) => {
                const odataProperty = this.metadata[column.referenceName];
                if (odataProperty == null) {
                    NotesService.instance.newNote('warning', `Field ${column.referenceName} in the select statement does not have an OData equivalent. Leaving it out of $select.`);
                } else {
                    if (!odataProperty.type.startsWith('Edm')) {
                        // Case: Complex type that has a default property (not expanded in the list of specialProperties).
                        const defaultField = ODataMetadataParser.defaultFields[odataProperty.type];
                        if (defaultField == null) {
                            NotesService.instance.newNote('warning', `The OData datatype of field ${column.referenceName} is not supported. Leaving it out of $select.`);
                        } else {
                            expandList.push(`${odataProperty.name}($select=${defaultField.defaultFieldName})`);
                        }
                    } else if (odataProperty.name.includes('/')) {
                        // Case: Complex type expanded in the list of specialProperties.
                        // Example: Area or Iteration.
                        const stringComponents = odataProperty.name.split('/');
                        expandList.push(`${stringComponents[0]}($select=${stringComponents[1]})`);
                    } else {
                        // Case: Primitive type.
                        selectList.push(odataProperty.name);
                    }
                }
            });

            return {
                select: (selectList.length === 0 ? null : `$select=${selectList.join(', ')}`),
                expand: (expandList.length === 0 ? null : `$expand=${expandList.join(', ')}`),
            };
        } else {
            NotesService.instance.newNote(
                'info',
                `every OData query should have a select statement. Please update your query to explicitly include a select statement.
                Without a select statement, the default and unexpanded will be returned by the query.`
            );
            return null;
        }
    }

    /**
     * @returns The $orderby statement, or null if it can't be parsed.
     */
    public makeOrderByStatement(): string {
        if (this.query.sortColumns != null) {
            const sortColumns = this.query.queryType === QueryType.Flat ? this.query.sortColumns : this.query.sortColumns.slice(this.query.sortColumns.length / 2);
            const sortStatements = sortColumns.map((column) => {
                const odataProperty = this.metadata[column.field.referenceName];
                if (odataProperty == null) {
                    NotesService.instance.newNote('warning', `Property ${column.field.referenceName} in the order by statement does not have an OData equivalent. Leaving it out of $orderby.`);
                } else if (!odataProperty.type.startsWith('Edm')) {
                    // Case: Complex type that has a default property (not expanded in the list of specialProperties).
                    const defaultField = ODataMetadataParser.defaultFields[odataProperty.type];
                    if (defaultField == null) {
                        NotesService.instance.newNote('warning', `The datatype of property ${column.field.referenceName} is not supported. Leaving it out of $orderby.`);
                    } else {
                        return `${odataProperty.name}/${defaultField.defaultFieldName} ${column.descending === true ? 'desc' : 'asc'}`;
                    }
                } else {
                    return `${odataProperty.name} ${column.descending === true ? 'desc' : 'asc'}`;
                }
                return null;
            });

            return sortStatements.filter((s) => s != null).length === 0 ? null : `$orderby=${sortStatements.filter((s) => s != null).join(', ')}`;
        }

        return null;
    }

    /**
     * @returns The $filter statement, or null if it can't be parsed.
     */
    public makeFilterStatement(rootClause: WorkItemQueryClause, superTypePrefix: string = ''): string {
        const oDataClauses = this.makeFilterStatementSimple(rootClause, superTypePrefix);

        if (oDataClauses == null) {
            return null;
        }

        // get ASOF. Unfortunately, the only way to currently detect ASOF is in the query text itself.
        const basicTokenizedQuery = this.query.wiql.toLocaleLowerCase().split(' ');
        const asofIndex = basicTokenizedQuery.findIndex((e) => e === 'asof');
        if (asofIndex !== -1) {
            // assuming thing after asof is a date. Assuming WIQL has been validated by ADO.
            const dateString = basicTokenizedQuery[asofIndex + 1];
            const isoDateString = new Date(dateString).toISOString();
            return `$filter=Revisions/any(r: ${oDataClauses} and (r/ChangedDate le ${isoDateString}) and (r/RevisedDate ge ${isoDateString} or r/RevisedDate eq null))`;
        }

        return `$filter=${oDataClauses}`;
    }

    /**
     * The filter statement, without ASOF and without $filter prepended.
     */
    public makeFilterStatementSimple(rootClause: WorkItemQueryClause, superTypePrefix: string = ''): string {
        if (rootClause == null) {
            return null;
        }

        const oDataClauses = this.visitClauses(rootClause, superTypePrefix);

        if (oDataClauses == null || oDataClauses === '' || oDataClauses === '()') {
            return null;
        }

        return oDataClauses;
    }

    /**
     * Recursively visits each clause.
     */
    private visitClauses(clause: WorkItemQueryClause, superTypePrefix: string): string {
        if (clause.logicalOperator != null) {
            // We have an array of clauses all joined by the logical operator.
            // Do DFS on the clauses.
            const clausesString = clause.clauses
                .map((c) => {
                    return this.visitClauses(c, superTypePrefix);
                })
                .filter((s) => s != null)
                .join(` ${LogicalOperation[clause.logicalOperator]} `); // 'and' or 'or'

            return clausesString.length === 0 ? null : `(${clausesString})`;
        } else {
            // We have a single clause. Parse it.
            const parsedClause = this.clauseParser.parseClause(clause);
            return parsedClause == null ? null : `${superTypePrefix != null && superTypePrefix !== '' ? `${superTypePrefix}/` : ``}${parsedClause}`;
        }
    }
}
