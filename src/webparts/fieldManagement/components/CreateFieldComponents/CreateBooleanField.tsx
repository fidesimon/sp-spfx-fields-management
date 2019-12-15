import * as React from 'react';
import { PrimaryButton, Button, Dropdown, IDropdownOption, TextField, Toggle } from 'office-ui-fabric-react';
import { FieldTypeKindEnum } from '../FieldTypeKindEnum';
import { ISPField } from '../SPField';

export interface CreateBooleanFieldProps {
    fieldTypeOptions: IDropdownOption[];
    saveButtonHandler: Function;
    cancelButtonHandler: Function;
    groupName: string;
}

export const CreateBooleanField: React.FC<CreateBooleanFieldProps> = (props) => {
    const [columnName, setColumnName] = React.useState("");
    const [fieldType, setFieldType] = React.useState(props.fieldTypeOptions[0].key);
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
            FieldTypeKind: FieldTypeKindEnum.Boolean,
            Group: group,
            Description: description,
            SchemaXml: `<Field Type="Boolean" Description="${description}" DisplayName="${columnName}" Required="FALSE" EnforceUniqueValues="FALSE" Group="${group}" StaticName="${internalName}" Name="${internalName}"><Default>${defaultValue}</Default></Field>`
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
                <Dropdown label="Field Type" options={props.fieldTypeOptions} defaultSelectedKey={fieldType} onChanged={(evt: any) => { setFieldType(evt.key)}} />
                <TextField label="Internal Name" required value={internalName} onChanged={(evt) => { setInternalName(evt)}} />
                <TextField label="Group" defaultValue={group} onChange={(evt: React.FormEvent<HTMLInputElement>) => { setGroup((evt.target as any).value);}} />
                <TextField label="Description" multiline autoAdjustHeight onChange={(evt: React.FormEvent<HTMLTextAreaElement>) => { setDescription((evt.target as any).value )}} />
                <Dropdown label="Default value" options={[{key: 1, text: 'Yes'}, {key: 0, text: 'No'}]} defaultSelectedKey={defaultValue} onChanged={(evt:any) => {
                    setDefaultValue(evt.key);
                }} />
                <br />
                <PrimaryButton text="Save" onClick={() => saveNewField()} />
                <Button text="Cancel" onClick={() => props.cancelButtonHandler()} />
            </>
    );
}