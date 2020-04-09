import * as React from 'react';
import { useEffect, useState } from 'react';
import { Dropdown, TextField, Toggle } from 'office-ui-fabric-react';
import { IPropertySchema } from '../../../models/IPropertySchema';

export interface IGenericObjectEditorProps {
    schema: any;
    object: any;
    defaultValues?: any;
    customRenderers?: any;
    ignoredProperties?: string[];
    readOnlyProperties?: string[];
    updateOnBlur?: boolean;
    fieldLabelGetter?: (field: string) => string;
    onObjectChanged?: (object: any) => void;
    children?: any;
}

export interface IPropertyPlaceholderProps { propertyName: string; }
export const PropertyPlaceholder = (props: IPropertyPlaceholderProps) => <></>;

export const GenericObjectEditor = (props: IGenericObjectEditorProps) => {

    const [objectProperties, setObjectProperties] = useState<string[]>([]);

    const getPropertyDefaultValueFromSchema = (propertyName: string) => {
        let propSchema = props.schema.properties[propertyName];
        if (propSchema) {
            switch (propSchema.type) {
                case 'string':
                    return '';
                case 'boolean':
                    return false;
                case 'number':
                    return 0;
                case 'object':
                    return {};
                default:
                    return null;
            }
        } else {
            return null;
        }
    };

    const getPropertyTypeFromSchema = (propertyName: string) => {
        let propSchema = props.schema.properties[propertyName];
        if (propSchema) {
            return propSchema.type;
        } else {
            return null;
        }
    };

    const isPropertyReadOnly = (propertyName: string) => {
        if (!props.readOnlyProperties || !props.readOnlyProperties.length) {
            return false;
        }

        return props.readOnlyProperties.indexOf(propertyName) > -1;
    };

    const onObjectPropertyChange = (propertyName: string, newValue: any) => {
        if (!props.onObjectChanged) {
            return;
        }

        let propertyType = getPropertyTypeFromSchema(propertyName);
        if (propertyType == 'number') {
            newValue = Number(newValue);
        }
        const updatedObject = { ...props.object, [propertyName]: newValue };

        // Set default values for properties of the argument object if not set
        objectProperties.forEach((p) => {
            // Get the property type

            let defaultValue =
                props.defaultValues && props.defaultValues[p]
                    ? props.defaultValues[p]
                    : getPropertyDefaultValueFromSchema(p);

            if (!updatedObject[p] && updatedObject[p] != false && updatedObject[p] != 0) {
                updatedObject[p] = defaultValue;
            }
        });

        props.onObjectChanged(updatedObject);
    };

    const getFieldLabel = (field: string, propertyDefinition?: IPropertySchema) => {
        if (props.fieldLabelGetter) {
            const foundLabel = props.fieldLabelGetter(field);
            if (foundLabel) {
                return foundLabel;
            }
        }

        // TODO Handle this from specified field label getter
        // Try translate from built-in resources
        // let key = 'PROP_' + field;
        // if (strings[key]) {
        //     return strings[key];
        // } else 
        if (propertyDefinition && propertyDefinition.title) {
            return propertyDefinition.title;
        } else {
            return field;
        }
    };

    const renderPropertyEditor = (propertyName: string, propertySchema: IPropertySchema) => {
        let { schema, customRenderers, defaultValues, object } = props;

        // Has custom renderer for the property
        if (customRenderers && customRenderers[propertyName]) {
            // If a default value is specified for current property and it is null, apply it
            if (!object[propertyName] && defaultValues && defaultValues[propertyName]) {
                object[propertyName] = defaultValues[propertyName];
            }

            return customRenderers[propertyName](object[propertyName]);
        }

        let isPropertyRequired =
            (schema.required && schema.required.length && schema.required.indexOf(propertyName) > -1) || false;

        if (propertySchema.enum) {
            if (propertySchema.enum.length > 1 || !isPropertyReadOnly(propertyName)) {
                return (
                    <Dropdown
                        required={isPropertyRequired}
                        label={getFieldLabel(propertyName, propertySchema)}
                        selectedKey={object[propertyName]}
                        options={propertySchema.enum.map((p) => ({ key: p, text: p }))}
                        onChanged={(value) => onObjectPropertyChange(propertyName, value.key)}
                    />
                );
            } else {
                return (
                    <TextField
                        label={getFieldLabel(propertyName, propertySchema)}
                        value={object[propertyName]}
                        readOnly={true}
                        required={isPropertyRequired}
                        onChange={(ev, value) => onObjectPropertyChange(propertyName, value)}
                    />
                );
            }
        } else {
            switch (propertySchema.type) {
                case 'boolean':
                    return (
                        <Toggle
                            label={getFieldLabel(propertyName, propertySchema)}
                            checked={object[propertyName] as boolean}
                            disabled={isPropertyReadOnly(propertyName)}
                            onChanged={(value) => onObjectPropertyChange(propertyName, value)}
                        />
                    );
                case 'array': // TODO Render a ArrayEditor for compl
                // case 'object': // TODO If object is a simple dictionary (key/non-complex object values) => Display a custom control
                // 	return (
                // 		<div>
                // 			<div className="ms-Grid-row">
                // 				<div className="ms-Grid-col ms-sm12">
                // 					<Label>{this._getFieldLabel(propertyName)}</Label>
                // 				</div>
                // 			</div>
                // 			<div className="ms-Grid-row">
                // 				<div className="ms-Grid-col ms-sm2">
                // 					<Icon iconName="InfoSolid" />
                // 				</div>
                // 				<div className="ms-Grid-col ms-sm10">
                // 					{strings.PropertyIsComplexTypeMessage}
                // 					<br />
                // 					{strings.UseJsonEditorMessage}
                // 				</div>
                // 			</div>
                // 		</div>
                // 	);
                case 'number':
                case 'string':
                default:
                    return (
                        <TextField
                            required={isPropertyRequired}
                            label={getFieldLabel(propertyName, propertySchema)}
                            value={object[propertyName]}
                            readOnly={isPropertyReadOnly(propertyName)}
                            onChange={(ev, value) => onObjectPropertyChange(propertyName, value)}
                        />
                    );
            }
        }
    };

    const refreshObjectProperties = () => {
        let { schema, ignoredProperties, defaultValues } = props;

        if (schema.type != 'object') {
            throw new Error('Cannot generate Object Editor from a non-object type');
        }

        if (!schema.properties || Object.keys(schema.properties).length == 0) {
            return;
        }

        let refreshedObjectProperties = Object.keys(schema.properties);
        if (ignoredProperties && ignoredProperties.length > 0) {
            refreshedObjectProperties = refreshedObjectProperties.filter((p) => ignoredProperties.indexOf(p) < 0);
        }
        setObjectProperties(refreshedObjectProperties);
    };

    // Use effects
    useEffect(() => {
        refreshObjectProperties();
    }, [props.schema]);

    // TODO See if really needed
    // private editTextValues: any;
    // private _onTextFieldValueChanged(fieldName: string, value: any) {
    //     if (this.props.updateOnBlur) {
    //         if (!this.editTextValues) {
    //             this.editTextValues = {};
    //         }
    //         this.editTextValues[fieldName] = value;
    //     } else {
    //         this._onObjectPropertyChange(fieldName, value);
    //     }
    // }

    // private _onTextFieldEdited(fieldName: string) {
    //     let value = this.editTextValues && this.editTextValues[fieldName];
    //     this._onObjectPropertyChange(fieldName, value);
    //     if (value) {
    //         delete this.editTextValues[fieldName];
    //     }
    // }

    const render = () => {
        let { schema, ignoredProperties, object } = props;
        if (!object) {
            return null;
        }

        let propertyEditors = {};
        objectProperties.forEach(p => {
            if (ignoredProperties && ignoredProperties.indexOf(p) >= 0) {
                return;
            }

            propertyEditors[p] = renderPropertyEditor(p, schema.properties[p]);
        });

        const renderChildrenRecursive = (node: any) => {
            return React.Children.map(node.props.children, (c, i) => {
                const asPlaceholder = c as React.ReactElement<IPropertyPlaceholderProps>;
                if (asPlaceholder.type == PropertyPlaceholder) {
                    const propName = asPlaceholder.props.propertyName;
                    return renderPropertyEditor(propName, schema.properties[propName]);
                } else {
                    if (React.Children.count(c.props.children) == 0) {
                        return c;
                    } else {
                        return React.cloneElement(c, { children: renderChildrenRecursive(c) });
                    }
                }
            });
        };

        if (React.Children.count(props.children) > 0) {
            return <>
                {renderChildrenRecursive(render())}
            </>;
        } else {
            return <>
                {Object.keys(propertyEditors).map(k => propertyEditors[k])}
            </>;
        }
    };

    return render();
};