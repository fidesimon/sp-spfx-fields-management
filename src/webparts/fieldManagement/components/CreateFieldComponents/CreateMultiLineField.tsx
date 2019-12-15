import * as React from 'react';
import { PrimaryButton, Button, Dropdown, IDropdownOption, TextField, Toggle } from 'office-ui-fabric-react';
import { FieldTypeKindEnum } from '../FieldTypeKindEnum';
import { ISPField } from '../SPField';

export interface CreateMultiLineFieldProps {
    fieldTypeOptions: IDropdownOption[];
    saveButtonHandler: Function;
    cancelButtonHandler: Function;
    groupName: string;
    onFieldTypeChange: Function;
}

export const CreateMultiLineField: React.FC<CreateMultiLineFieldProps> = (props) => {
    const [columnName, setColumnName] = React.useState("");
    const [fieldType, setFieldType] = React.useState(FieldTypeKindEnum.Note);
    const [internalName, setInternalName] = React.useState("");
    const [group, setGroup] = React.useState(props.groupName);
    const [description, setDescription] = React.useState("");
    const [required,setRequired] = React.useState(false);
    const [allowUnlimitedLength, setAllowUnlimitedLength] = React.useState(false);
    const [noOfLines, setNoOfLines] = React.useState(6);
    const [allowEnhancedRichText, setAllowEnhancedRichText] = React.useState(true);
    const [appendChangesToExistingText, setAppendChangesToExistingText] = React.useState(false);

    const saveNewField = () => {
        let body: ISPField;
        body = {
            "@odata.type": "#SP.FieldMultiLineText",
            Title: columnName,
            StaticName: internalName,
            InternalName: internalName,
            FieldTypeKind: FieldTypeKindEnum.Note,
            Required: required,
            Group: group,
            Description: description,

            UnlimitedLengthInDocumentLibrary: allowUnlimitedLength,
            AppendOnly: appendChangesToExistingText,
            NumberOfLines: noOfLines,
            RichText: allowEnhancedRichText,
            SchemaXml: `<Field
                        Name="${internalName}"
                        DisplayName="${columnName}"
                        Description="${description}"
                        StaticName="${internalName}"
                        Group="${group}"
                        Type="Note"
                        NumLines="${noOfLines}"
                        UnlimitedLengthInDocumentLibrary="${(allowUnlimitedLength ? "TRUE" : "FALSE")}" 
                        Required="${(required ? "TRUE" : "FALSE")}" 
                        AppendOnly="${(appendChangesToExistingText ? "TRUE" : "FALSE")}" 
                        RichText="${(allowEnhancedRichText ? "TRUE" : "FALSE")}"
                        />`
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
                <Dropdown label="Field Type" options={props.fieldTypeOptions} defaultSelectedKey={fieldType} onChanged={(evt: any) => { props.onFieldTypeChange(evt.key)}} />
                <TextField label="Internal Name" required value={internalName} onChanged={(evt) => { setInternalName(evt)}} />
                <TextField label="Group" defaultValue={group} onChange={(evt: React.FormEvent<HTMLInputElement>) => { setGroup((evt.target as any).value);}} />
                <TextField label="Description" multiline autoAdjustHeight onChange={(evt: React.FormEvent<HTMLTextAreaElement>) => { setDescription((evt.target as any).value )}} />
                <Toggle label="Required" onChanged={(evt) => setRequired(evt)} />
                <Toggle label="Allow unlimited length in document libraries" onChanged={(evt) => setAllowUnlimitedLength(evt)} />
                <TextField label="Number of lines for editing" max={255} min={0} type="number" defaultValue={noOfLines.toString()} onChange={(evt: React.FormEvent<HTMLInputElement>) => { setNoOfLines(+((evt.target as any).value));}} />
                <Toggle label="Allow enhanced rich text" checked={allowEnhancedRichText} onChanged={(evt) => {
                    setAllowEnhancedRichText(evt);}
                    } /> 
                <Toggle label="Append Changes to Existing Text" onChanged={(evt) => setAppendChangesToExistingText(evt)} />
                <br />
                <PrimaryButton text="Save" onClick={() => saveNewField()} />
                <Button text="Cancel" onClick={() => props.cancelButtonHandler()} />
            </>
    );
}