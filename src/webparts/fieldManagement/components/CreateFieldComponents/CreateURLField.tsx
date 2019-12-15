import * as React from 'react';
import { PrimaryButton, Button, Dropdown, IDropdownOption, TextField, Toggle } from 'office-ui-fabric-react';
import { FieldTypeKindEnum } from '../FieldTypeKindEnum';
import { ISPField } from '../SPField';
import { ICreateFieldProps } from './ICreateFieldProps';

export const CreateURLField: React.FC<ICreateFieldProps> = (props) => {
    const [columnName, setColumnName] = React.useState("");
    const fieldType = FieldTypeKindEnum.URL;
    const [internalName, setInternalName] = React.useState("");
    const [group, setGroup] = React.useState(props.groupName);
    const [description, setDescription] = React.useState("");
    const [required, setRequired] = React.useState(false);
    const [urlFieldFormat, setUrlFieldFormat] = React.useState("Hyperlink");

    const saveNewField = () => {
        let body: ISPField;
        body = {
            "@odata.type": "#SP.FieldUrl",
            Title: columnName,
            StaticName: internalName,
            InternalName: internalName,
            FieldTypeKind: FieldTypeKindEnum.URL,
            Required: required,
            Group: group,
            Description: description,
            SchemaXml: `<Field Type="URL" Description="${description}" Format="${urlFieldFormat}" DisplayName="${columnName}" Required="${required ? 'TRUE' : 'FALSE'}" EnforceUniqueValues="FALSE" Group="${group}" StaticName="${internalName}" Name="${internalName}"></Field>`
        };

        props.saveButtonHandler(body);
    }

    const urlFormatOptions: IDropdownOption[] = [
        {key: 'Hyperlink', text: 'Hyperlink'}, 
        {key: 'Image', text: 'Picture'}
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
                <Toggle label="Required" onChanged={(evt) => setRequired(evt)} />
                <Dropdown label="Format URL as" options={urlFormatOptions} defaultSelectedKey={urlFieldFormat} onChanged={(evt:any) => {
                    setUrlFieldFormat(evt.key);
                }} />
                <br />
                <PrimaryButton text="Save" onClick={() => saveNewField()} />
                <Button text="Cancel" onClick={() => props.cancelButtonHandler()} />
            </>
    );
}