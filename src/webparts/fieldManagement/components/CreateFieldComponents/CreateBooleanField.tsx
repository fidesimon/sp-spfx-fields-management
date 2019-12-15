import * as React from 'react';
import { PrimaryButton, Button, Dropdown, TextField, IDropdownOption } from 'office-ui-fabric-react';
import { FieldTypeKindEnum } from '../FieldTypeKindEnum';
import { ISPField } from '../SPField';
import { ICreateFieldProps } from './ICreateFieldProps';

export const CreateBooleanField: React.FC<ICreateFieldProps> = (props) => {
    const [columnName, setColumnName] = React.useState("");
    const fieldType = FieldTypeKindEnum.Boolean;
    const [internalName, setInternalName] = React.useState("");
    const [group, setGroup] = React.useState(props.groupName);
    const [description, setDescription] = React.useState("");
    const [defaultValue, setDefaultValue] = React.useState<number>();

    const saveNewField = () => {
        let body: ISPField;
        body = {
            "@odata.type": "#SP.Field",
            Title: columnName,
            StaticName: internalName,
            InternalName: internalName,
            FieldTypeKind: fieldType,
            Group: group,
            Description: description,
            SchemaXml: `<Field Type="Boolean" Description="${description}" DisplayName="${columnName}" Required="FALSE" EnforceUniqueValues="FALSE" Group="${group}" StaticName="${internalName}" Name="${internalName}"><Default>${defaultValue}</Default></Field>`
        };

        props.saveButtonHandler(body);
    }

    const defaultValueOptions: IDropdownOption[] = [
        {key: 1, text: 'Yes'}, 
        {key: 0, text: 'No'}
    ];

    return (
            <>
                <TextField label="Column Name" required value={columnName} onChanged={(evt) => {
                    setColumnName(evt);
                    let newValue = evt.replace(/[^A-Z0-9]+/ig, "");
                    setInternalName(newValue.length >= 32 ? newValue.substr(0, 32) : newValue);
                }} />
                <Dropdown label="Field Type" options={props.fieldTypeOptions} defaultSelectedKey={fieldType} onChanged={(evt: any) => { props.onFieldTypeChange(evt.key)}} />
                <TextField label="Internal Name" required value={internalName} onChanged={(evt) => { setInternalName(evt)}} />
                <TextField label="Group" defaultValue={group} onChange={(evt: React.FormEvent<HTMLInputElement>) => { setGroup((evt.target as any).value);}} />
                <TextField label="Description" multiline autoAdjustHeight onChange={(evt: React.FormEvent<HTMLTextAreaElement>) => { setDescription((evt.target as any).value )}} />
                <Dropdown label="Default value" options={defaultValueOptions} defaultSelectedKey={defaultValue} onChanged={(evt:any) => {
                    setDefaultValue(evt.key);
                }} />
                <br />
                <PrimaryButton text="Save" onClick={() => saveNewField()} />
                <Button text="Cancel" onClick={() => props.cancelButtonHandler()} />
            </>
    );
}