import * as React from 'react';
import { PrimaryButton, Button, Dropdown, TextField, Toggle, IDropdownOption } from 'office-ui-fabric-react';
import { FieldTypeKindEnum } from '../FieldTypeKindEnum';
import { ISPField } from '../SPField';
import { ICreateFieldProps } from './ICreateFieldProps';
import { BaseComponentContext } from '@microsoft/sp-component-base';

interface ICreateLookupFieldProps extends ICreateFieldProps{
    context: BaseComponentContext;
}

export const CreateLookupField: React.FC<ICreateLookupFieldProps> = (props) => {
    const [columnName, setColumnName] = React.useState("");
    const fieldType = FieldTypeKindEnum.User;
    const [internalName, setInternalName] = React.useState("");
    const [group, setGroup] = React.useState(props.groupName);
    const [description, setDescription] = React.useState("");
    const [required, setRequired] = React.useState(false);
    const [enforceUniqueValues, setEnforceUniqueValues] = React.useState<boolean>(false);

    const saveNewField = () => {
        let body: ISPField;

        body = {
            "@odata.type": "#SP.FieldLookup",
            Title: columnName,
            StaticName: internalName,
            InternalName: internalName,
            FieldTypeKind: fieldType,
            Required: required,
            EnforceUniqueValues: enforceUniqueValues,
            Group: group,
            Description: description,
            SchemaXml: `<Field Type="Lookup" Description="${description}" DisplayName="${columnName}" Required="${(required ? "TRUE" : "FALSE")}" EnforceUniqueValues="${(enforceUniqueValues ? "TRUE" : "FALSE")}" Group="${group}" StaticName="${internalName}" Name="${internalName}"></Field>`
        };

        props.saveButtonHandler(body);
    }

    return (
        <>
            <TextField label="Column Name" required value={columnName} onChanged={(evt) => {
                setColumnName(evt);
                let newValue = evt.replace(/[^A-Z0-9]+/ig, "");
                setInternalName(newValue.length >= 32 ? newValue.substr(0, 32) : newValue);
            }} />
            <Dropdown label="Field Type" options={props.fieldTypeOptions} defaultSelectedKey={fieldType} onChanged={(evt: any) => { props.onFieldTypeChange(evt.key) }} />
            <TextField label="Internal Name" required value={internalName} onChanged={(evt) => { setInternalName(evt) }} />
            <TextField label="Group" defaultValue={group} onChange={(evt: React.FormEvent<HTMLInputElement>) => { setGroup((evt.target as any).value); }} />
            <TextField label="Description" multiline autoAdjustHeight onChange={(evt: React.FormEvent<HTMLTextAreaElement>) => { setDescription((evt.target as any).value) }} />
            <Toggle label="Required" onChanged={(evt) => setRequired(evt)} />
            <Toggle label="Enforce Unique Values" onChanged={(evt) => setEnforceUniqueValues(evt)} />
            <br />
            <PrimaryButton text="Save" onClick={() => saveNewField()} />
            <Button text="Cancel" onClick={() => props.cancelButtonHandler()} />
        </>
    );
}