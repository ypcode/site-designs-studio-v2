export interface IPropertySchema {
    type?: string;
    enum?: string[];
    title?: string;
    description?: string;
    properties?: { [property: string]: IPropertySchema };
}