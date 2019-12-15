import * as React from 'react';
import { PrimaryButton, Button, Dropdown, IDropdownOption, TextField, Toggle, ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react';
import { FieldTypeKindEnum } from '../FieldTypeKindEnum';
import { ISPField } from '../SPField';
import { ICreateFieldProps } from './ICreateFieldProps';

export const CreateChoiceField: React.FC<ICreateFieldProps> = (props) => {
    const [columnName, setColumnName] = React.useState("");
    const [fieldType, setFieldType] = React.useState(FieldTypeKindEnum.Choice);
    const [internalName, setInternalName] = React.useState("");
    const [group, setGroup] = React.useState(props.groupName);
    const [description, setDescription] = React.useState("");
    const [required,setRequired] = React.useState(false);
    const [enforceUniqueValues, setEnforceUniqueValues] = React.useState<boolean>(false);
    const [defaultValue, setDefaultValue] = React.useState();
    const [choices, setChoices] = React.useState(["Enter Choice #1", "Enter Choice #2", "Enter Choice #3"]);
    const [choiceFormat, setChoiceFormat] = React.useState('Dropdown');
    const [fillInChoice, setFillInChoice] = React.useState(false);
    const [defaultValueChoices, setDefaultValueChoices] = React.useState<IDropdownOption[]>([{key: '', text: '(empty)', isSelected: true}, {key: 'Enter Choice #1', text: 'Enter Choice #1'}, {key: 'Enter Choice #2', text: 'Enter Choice #2'},{key: 'Enter Choice #3', text: 'Enter Choice #3'}]);

    const saveNewField = () => {
        let body: ISPField;
        let choicesString = `<CHOICES><CHOICE>${choices.join("</CHOICE><CHOICE>")}</CHOICE></CHOICES>`;
        let defaultString: string = (defaultValue == null || defaultValue.length == '') ? '' : `<Default>${defaultValue}</Default>`;
        body = {
            "@odata.type": "#SP.FieldChoice",
            Title: columnName,
            StaticName: internalName,
            InternalName: internalName,
            FieldTypeKind: FieldTypeKindEnum.Choice,
            Required: required,
            EnforceUniqueValues: enforceUniqueValues,
            DefaultValue: defaultValue,
            Group: group,
            Description: description,
            SchemaXml: `<Field Type="Choice" Description="${description}" DisplayName="${columnName}" Format="${choiceFormat}" FillInChoice="${(fillInChoice ? "TRUE" : "FALSE")}" Required="${(required? "TRUE" : "FALSE")}" EnforceUniqueValues="${(enforceUniqueValues? "TRUE" : "FALSE")}" Group="${group}" StaticName="${internalName}" Name="${internalName}">${defaultString}${choicesString}</Field>`
        };

        props.saveButtonHandler(body);
    }

    const choiceFieldFormatOptions: IChoiceGroupOption[] = [
        {
            key: 'Dropdown',
            text: 'Drop-Down Menu',
        },
        {
            key: 'RadioButtons',
            text: 'Radio Buttons'
        },
        {
            key: 'CheckBoxes',
            text: 'Checkboxes (allow multiple selection)',
            disabled: true
        }
    ];

    const distinct = (value, index, self) => {
        return self.indexOf(value) === index;
    };

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
                <Toggle label="Enforce Unique Values" onChanged={(evt) => setEnforceUniqueValues(evt)} />
                <TextField label="Type each choice on a separate line" 
                    defaultValue={`Enter Choice #1
Enter Choice #2
Enter Choice #3`} 
                    multiline 
                    autoAdjustHeight 
                    onChange={(choices: React.FormEvent<HTMLTextAreaElement>) => { 
                            let distinctChoices = (choices.target as any).value.split('\n').filter(n => n!= '').filter(distinct);
                            let defaultValueChoices: IDropdownOption[] = distinctChoices.map((item)=>{
                                return {key: item, text: item};
                            });
                            defaultValueChoices.unshift({key: '', text: '(empty)', isSelected: true});
                            setChoices(distinctChoices);
                            setDefaultValueChoices(defaultValueChoices);
                        } 
                    }
                />
                <Dropdown label="Default value" defaultValue="(empty)" options={defaultValueChoices} onChanged={(evt: any) => {
                    setDefaultValue(evt.key);
                }} />
                <ChoiceGroup label="Display choices using" defaultSelectedKey={choiceFormat} options={choiceFieldFormatOptions} onChanged={(evt: any) => { 
                    setChoiceFormat(evt.key);
                }} />
                <Toggle label="Allow 'Fill-in' choices" onChanged={(evt) => setFillInChoice(evt)} />
                <br />
                <PrimaryButton text="Save" onClick={() => saveNewField()} />
                <Button text="Cancel" onClick={() => props.cancelButtonHandler()} />
            </>
    );
}