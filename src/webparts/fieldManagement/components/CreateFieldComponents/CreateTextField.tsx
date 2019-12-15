import * as React from 'react';
import { PrimaryButton, Button, Dropdown, IDropdownOption, TextField, Toggle } from 'office-ui-fabric-react';
import { FieldTypeKindEnum } from '../FieldTypeKindEnum';
import { ISPField } from '../SPField';

export interface CreateTextFieldProps {
    fieldTypeOptions: IDropdownOption[];
    saveButtonHandler: Function;
    cancelButtonHandler: Function;
    groupName: string;
}

export const CreateTextField: React.FC<CreateTextFieldProps> = (props) => {
    const [columnName, setColumnName] = React.useState("");
    const [fieldType, setFieldType] = React.useState(props.fieldTypeOptions[0].key);
    const [internalName, setInternalName] = React.useState("");
    const [group, setGroup] = React.useState(props.groupName);
    const [description, setDescription] = React.useState("");
    const [required,setRequired] = React.useState(false);
    const [enforceUniqueValues, setEnforceUniqueValues] = React.useState<boolean>(false);
    const [maxNoOfChars, setMaxNoOfChars] = React.useState<number>(255);
    const [defaultValue, setDefaultValue] = React.useState();

    const saveNewField = () => {
        let body: ISPField;
        let defaultString: string = (defaultValue == undefined || defaultValue.length == 0) ? '' : "<Default>" + defaultValue + "</Default>";
        body = {
            "@odata.type": "#SP.FieldText",
            Title: columnName,
            StaticName: internalName,
            InternalName: internalName,
            FieldTypeKind: FieldTypeKindEnum.Text,
            Required: required,
            EnforceUniqueValues: enforceUniqueValues,
            MaxLength: maxNoOfChars,
            DefaultValue: defaultValue,
            Group: group,
            Description: description,
            SchemaXml: `<Field Type="Text" Description="${description}" DisplayName="${columnName}" Required="${(required? "TRUE" : "FALSE")}" EnforceUniqueValues="${(enforceUniqueValues? "TRUE" : "FALSE")}" Group="${group}" StaticName="${internalName}" Name="${internalName}">${defaultString}</Field>`
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
                <Toggle label="Required" onChanged={(evt) => setRequired(evt)} />
                <Toggle label="Enforce Unique Values" onChanged={(evt) => setEnforceUniqueValues(evt)} />
                <TextField label="Maximum number of characters" max={255} min={0} type="number" defaultValue={maxNoOfChars.toString()} onChange={(evt: React.FormEvent<HTMLInputElement>) => { setMaxNoOfChars(+((evt.target as any).value));}} />
                <TextField label="Default value" onChange={(evt: React.FormEvent<HTMLInputElement>) => { setDefaultValue((evt.target as any).value);}} />
                <br />
                <PrimaryButton text="Save" onClick={() => saveNewField()} />
                <Button text="Cancel" onClick={() => props.cancelButtonHandler()} />
            </>
    );
}