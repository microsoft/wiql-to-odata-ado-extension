/**
 * Copyright (c) Microsoft Corporation.  All rights reserved.
 */

export interface IEntityType {
    name: string;
    key?: IPropertyRef[];
    properties?: IProperty[];
}

export interface IPropertyRef {
    name: string;
}

export interface IProperty {
    name: string;
    type: string;
    nullable?: boolean;
    annotations?: IAnnotation[];
}

export interface IReferentialConstraint {
    property: string;
    referencedProperty: string;
}

export interface IEntityContainer {
    name: string;
    entitySets?: IEntitySet[];
}

export interface IEntitySet {
    name: string;
    entityType: string;
    navigationPropertyBindings?: INavigationPropertyBinding[];
    annotations?: IAnnotation[];
}

export interface INavigationPropertyBinding {
    path: string;
    target: string;
}

export interface IAnnotation {
    term: string;
    value: any;
}

export interface ISchema {
    namespace: string;
    entityTypes?: IEntityType[];
    entityContainers?: IEntityContainer[];
}

export interface IMetadata {
    schemas?: ISchema[];
}

export interface IDefaultField {
    fieldType: string;
    defaultFieldName: string;
    defaultFieldType: string;
}

/**
 * Properties that for some reason do not have a ReferenceName. These are added to the property list when parsing.
 * Some, such as Iteration and Area, are premapped to primative field,
 * while others are mapped to a complex type that must be used in conjunction with defaultFields
 */
export const specialProperties: IProperty[] = [
    { name: 'Project', type: 'Microsoft.VisualStudio.Services.Analytics.Model.Project', annotations: [{ term: 'Ref.ReferenceName', value: 'System.TeamProject' }] },
    { name: 'Iteration/IterationId', type: 'Edm.Guid', annotations: [{ term: 'Ref.ReferenceName', value: 'System.IterationId' }] },
    { name: 'Iteration/IterationPath', type: 'Edm.String', annotations: [{ term: 'Ref.ReferenceName', value: 'System.IterationPath' }] },
    { name: 'Area/AreaId', type: 'Edm.Guid', annotations: [{ term: 'Ref.ReferenceName', value: 'System.AreaId' }] },
    { name: 'Area/AreaPath', type: 'Edm.String', annotations: [{ term: 'Ref.ReferenceName', value: 'System.AreaPath' }] },
    // Below only used in Links expand, so don't need to include 'Links/'...
    { name: 'LinkTypeReferenceName', type: 'Edm.String', annotations: [{ term: 'Ref.ReferenceName', value: 'System.Links.LinkType' }] },
    { name: 'ChangedOn', type: 'Microsoft.VisualStudio.Services.Analytics.Model.CalendarDate', annotations: [{ term: 'Ref.ReferenceName', value: 'System.ChangedDate' }] },
    { name: 'ClosedOn', type: 'Microsoft.VisualStudio.Services.Analytics.Model.CalendarDate', annotations: [{ term: 'Ref.ReferenceName', value: 'Microsoft.VSTS.Common.ClosedDate' }] },
    { name: 'CreatedOn', type: 'Microsoft.VisualStudio.Services.Analytics.Model.CalendarDate', annotations: [{ term: 'Ref.ReferenceName', value: 'System.CreatedDate' }] },
    { name: 'ResolvedOn', type: 'Microsoft.VisualStudio.Services.Analytics.Model.CalendarDate', annotations: [{ term: 'Ref.ReferenceName', value: 'Microsoft.VSTS.Common.ResolvedDate' }] },
    { name: 'StateChangeOn', type: 'Microsoft.VisualStudio.Services.Analytics.Model.CalendarDate', annotations: [{ term: 'Ref.ReferenceName', value: 'Microsoft.VSTS.Common.StateChangeDate' }] },
    { name: 'Iteration/IterationLevel1', type: 'Edm.String', annotations: [{ term: 'Ref.ReferenceName', value: 'System.IterationLevel1' }] },
    { name: 'Iteration/IterationLevel2', type: 'Edm.String', annotations: [{ term: 'Ref.ReferenceName', value: 'System.IterationLevel2' }] },
    { name: 'Iteration/IterationLevel3', type: 'Edm.String', annotations: [{ term: 'Ref.ReferenceName', value: 'System.IterationLevel3' }] },
    { name: 'Iteration/IterationLevel4', type: 'Edm.String', annotations: [{ term: 'Ref.ReferenceName', value: 'System.IterationLevel4' }] },
    { name: 'Iteration/IterationLevel5', type: 'Edm.String', annotations: [{ term: 'Ref.ReferenceName', value: 'System.IterationLevel5' }] },
    { name: 'Iteration/IterationLevel6', type: 'Edm.String', annotations: [{ term: 'Ref.ReferenceName', value: 'System.IterationLevel6' }] },
    { name: 'Iteration/IterationLevel7', type: 'Edm.String', annotations: [{ term: 'Ref.ReferenceName', value: 'System.IterationLevel7' }] },
    { name: 'Iteration/IterationLevel8', type: 'Edm.String', annotations: [{ term: 'Ref.ReferenceName', value: 'System.IterationLevel8' }] },
    { name: 'Iteration/IterationLevel9', type: 'Edm.String', annotations: [{ term: 'Ref.ReferenceName', value: 'System.IterationLevel9' }] },
    { name: 'Iteration/IterationLevel10', type: 'Edm.String', annotations: [{ term: 'Ref.ReferenceName', value: 'System.IterationLevel10' }] },
    { name: 'Iteration/IterationLevel11', type: 'Edm.String', annotations: [{ term: 'Ref.ReferenceName', value: 'System.IterationLevel11' }] },
    { name: 'Iteration/IterationLevel12', type: 'Edm.String', annotations: [{ term: 'Ref.ReferenceName', value: 'System.IterationLevel12' }] },
    { name: 'Iteration/IterationLevel13', type: 'Edm.String', annotations: [{ term: 'Ref.ReferenceName', value: 'System.IterationLevel13' }] },
    { name: 'Iteration/IterationLevel14', type: 'Edm.String', annotations: [{ term: 'Ref.ReferenceName', value: 'System.IterationLevel14' }] },
    { name: 'Area/AreaLevel1', type: 'Edm.String', annotations: [{ term: 'Ref.ReferenceName', value: 'System.AreaLevel1' }] },
    { name: 'Area/AreaLevel2', type: 'Edm.String', annotations: [{ term: 'Ref.ReferenceName', value: 'System.AreaLevel2' }] },
    { name: 'Area/AreaLevel3', type: 'Edm.String', annotations: [{ term: 'Ref.ReferenceName', value: 'System.AreaLevel3' }] },
    { name: 'Area/AreaLevel4', type: 'Edm.String', annotations: [{ term: 'Ref.ReferenceName', value: 'System.AreaLevel4' }] },
    { name: 'Area/AreaLevel5', type: 'Edm.String', annotations: [{ term: 'Ref.ReferenceName', value: 'System.AreaLevel5' }] },
    { name: 'Area/AreaLevel6', type: 'Edm.String', annotations: [{ term: 'Ref.ReferenceName', value: 'System.AreaLevel6' }] },
    { name: 'Area/AreaLevel7', type: 'Edm.String', annotations: [{ term: 'Ref.ReferenceName', value: 'System.AreaLevel7' }] },
    { name: 'Area/AreaLevel8', type: 'Edm.String', annotations: [{ term: 'Ref.ReferenceName', value: 'System.AreaLevel8' }] },
    { name: 'Area/AreaLevel9', type: 'Edm.String', annotations: [{ term: 'Ref.ReferenceName', value: 'System.AreaLevel9' }] },
    { name: 'Area/AreaLevel10', type: 'Edm.String', annotations: [{ term: 'Ref.ReferenceName', value: 'System.AreaLevel10' }] },
    { name: 'Area/AreaLevel11', type: 'Edm.String', annotations: [{ term: 'Ref.ReferenceName', value: 'System.AreaLevel11' }] },
    { name: 'Area/AreaLevel12', type: 'Edm.String', annotations: [{ term: 'Ref.ReferenceName', value: 'System.AreaLevel12' }] },
    { name: 'Area/AreaLevel13', type: 'Edm.String', annotations: [{ term: 'Ref.ReferenceName', value: 'System.AreaLevel13' }] },
    { name: 'Area/AreaLevel14', type: 'Edm.String', annotations: [{ term: 'Ref.ReferenceName', value: 'System.AreaLevel14' }] },
];

/**
 * A list of complex types that can be simplified to primitive types given a default property.
 */
export const defaultFields: IDefaultField[] =  [
    { fieldType: 'Microsoft.VisualStudio.Services.Analytics.Model.CalendarDate', defaultFieldName: 'Date', defaultFieldType: 'Edm.DateTimeOffset' },
    { fieldType: 'Microsoft.VisualStudio.Services.Analytics.Model.Project', defaultFieldName: 'ProjectName', defaultFieldType: 'Edm.String' },
    { fieldType: 'Microsoft.VisualStudio.Services.Analytics.Model.User', defaultFieldName: 'UserEmail', defaultFieldType: 'Edm.String' },
];

/**
 * Manages the metadata xml containing all OData fields and types.
 */
export class ODataMetadataParser {
    private static defaultFieldsInternal = null;

    /**
     * Parses the $metadata xml into an object.
     * @param document The xml returned from the odata $metadata endpoint
     */
    public static parseDocument(document: XMLDocument): IMetadata {
        const edmx = document.getElementsByTagName('edmx:Edmx')[0];
        return {
            schemas: this.parseCollection(edmx, 'Schema', (e) => this.parseSchema(e)),
        };
    }

    /**
     * @returns A dictionary mapping Ref.ReferenceNames to IProperties.
     * @param metadata IMetadata to convert to a dictionary.
     */
    public static createPropertyMap(metadata: IMetadata): { [id: string]: IProperty; } {
        return metadata.schemas
            .reduce<IEntityType[]>(
                (acc, schema) => (schema.entityTypes ? acc.concat(schema.entityTypes) : acc) as IEntityType[],
                new Array<IEntityType>()
            )
            .filter((entityType) => {
                return entityType.name === 'WorkItem' || entityType.name === 'CustomWorkItem';
            })
            .reduce<IProperty[]>(
                (acc, x) => (x.properties ? acc.concat(x.properties) : acc) as IProperty[],
                new Array<IProperty>()
            )
            .map((p) => {
                const referenceName = p.annotations
                    ? p.annotations
                        .filter((a) => a.term === 'Ref.ReferenceName')
                        .map((a) => a.value as string)[0]
                    : undefined;

                return { referenceName, property: p };
            })
            .filter((x) => {
                return x.referenceName !== undefined;
            })
            .reduce<{ [id: string]: IProperty; }>(
                (acc, x) => { acc[x.referenceName] = x.property; return acc; },
                {}
            );
    }

    /**
     * @returns a mapping of OData type names to objects defining the field that should be used to simplify them to primative types.
     */
    public static get defaultFields(): { [name: string]: IDefaultField; } {
        if (ODataMetadataParser.defaultFieldsInternal == null) {
            ODataMetadataParser.defaultFieldsInternal = defaultFields.reduce<{ [name: string]: IDefaultField; }>(
                (acc, x) => { acc[x.fieldType] = x; return acc; },
                {}
            );
        }
        return ODataMetadataParser.defaultFieldsInternal;
    }

    //#region "Xml Parsers"

    private static parseSchema(element: Element): ISchema {
        return {
            namespace: element.getAttribute('Namespace'),
            entityTypes: this.parseCollection<IEntityType>(element, 'EntityType', (e) => this.parseEntityType(e)),
            entityContainers: this.parseCollection<IEntityContainer>(element, 'EntityContainer', (e) => this.parseEntityContainer(e)),
        };
    }

    private static parseEntityContainer(element: Element): IEntityContainer {
        return {
            name: element.getAttribute('Name'),
            entitySets: this.parseCollection<IEntitySet>(element, 'EntitySet', (e) => this.parseEntitySet(e)),
        };
    }

    private static parseEntitySet(element: Element): IEntitySet {
        return {
            name: element.getAttribute('Name'),
            entityType: element.getAttribute('EntityType'),
            navigationPropertyBindings: this.parseCollection<INavigationPropertyBinding>(
                element,
                'NavigationPropertyBinding',
                (e) => this.parseNavigationPropertyBinding(e)
            ),
            annotations: this.parseCollection(element, 'Annotation', (e) => this.parseAnnotation(e)),
        };
    }

    private static parseNavigationPropertyBinding(element: Element): INavigationPropertyBinding {
        return {
            path: element.getAttribute('Path'),
            target: element.getAttribute('Target'),
        };
    }

    private static parseEntityType(element: Element): IEntityType {
        return {
            name: element.getAttribute('Name'),
            properties: [
                ...this.parseCollection<IProperty>(element, 'Property', (e) => this.parseProperty(e)),
                ...this.parseCollection<IProperty>(element, 'NavigationProperty', (e) => this.parseProperty(e)),
                ...specialProperties,
            ],
        };
    }

    private static parseProperty(element: Element): IProperty {
        return {
            name: element.getAttribute('Name'),
            type: element.getAttribute('Type'),
            nullable: element.hasAttribute('Nullable') ? !!element.getAttribute('Nullable') : undefined,
            annotations: this.parseCollection(element, 'Annotation', (e) => this.parseAnnotation(e)),
        };
    }

    private static parseAnnotation(element: Element): IAnnotation {
        // TODO: support different types of annotations
        return {
            term: element.getAttribute('Term'),
            value: element.getAttribute('String'),
        };
    }

    private static parseCollection<T>(element: Element, name: string, select: (element: Element) => T) {
        const nodes = element.getElementsByTagName(name);
        const result = Array.from(nodes).map((node) => select(node));

        return result.length > 0 ? result : undefined;
    }

    //#endregion "Xml Parsers"
}
