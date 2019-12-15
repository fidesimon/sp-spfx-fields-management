import * as React from 'react';
import { PrimaryButton, Button, Dropdown, IDropdownOption, TextField, Toggle } from 'office-ui-fabric-react';
import { FieldTypeKindEnum } from '../FieldTypeKindEnum';
import { ISPField } from '../SPField';

export interface CreateNumberFieldProps {
    fieldTypeOptions: IDropdownOption[];
    saveButtonHandler: Function;
    cancelButtonHandler: Function;
    groupName: string;
}

export const CreateNumberField: React.FC<CreateNumberFieldProps> = (props) => {
    const [columnName, setColumnName] = React.useState("");
    const [fieldType, setFieldType] = React.useState(props.fieldTypeOptions[0].key);
    const [internalName, setInternalName] = React.useState("");
    const [group, setGroup] = React.useState(props.groupName);
    const [description, setDescription] = React.useState("");
    const [required,setRequired] = React.useState(false);
    const [enforceUniqueValues, setEnforceUniqueValues] = React.useState<boolean>(false);
    const [defaultValue, setDefaultValue] = React.useState();
    const [showAsPercentage, setShowAsPercentage] = React.useState(false);
    const [minValue, setMinValue] = React.useState();
    const [maxValue, setMaxValue] = React.useState();
    const [displayFormat, setDisplayFormat] = React.useState(-1);

    const saveNewField = () => {
        let body: ISPField;
        let defaultString = defaultValue.length == 0 ? '' : "<Default>" + (showAsPercentage ? (+(defaultValue)/100).toString() : defaultValue) + "</Default>";
        let minString = minValue == null ? '' : (showAsPercentage ? 'Min="' + minValue/100 + '"' : 'Min="' + minValue + '"');
        let maxString = maxValue == null ? '' : (showAsPercentage ? 'Max="' + maxValue/100 + '"' : 'Max="' + maxValue + '"');
        
        body = {
            "@odata.type": "#SP.FieldNumber",
            Title: columnName,
            StaticName: internalName,
            InternalName: internalName,
            FieldTypeKind: FieldTypeKindEnum.Number,
            Required: required,
            EnforceUniqueValues: enforceUniqueValues,
            DefaultValue: defaultValue,
            Group: group,
            DisplayFormat: +(displayFormat),
            ShowAsPercentage: showAsPercentage,
            Description: description,
            SchemaXml: `<Field Type="Number" 
            Description="${description}" 
            DisplayName="${columnName}" 
            Required="${(required? "TRUE" : "FALSE")}" 
            EnforceUniqueValues="${(enforceUniqueValues? "TRUE" : "FALSE")}" 
            Group="${group}" 
            StaticName="${internalName}" 
            Percentage="${(showAsPercentage? "TRUE" : "FALSE")}" 
            Decimals="${displayFormat}" 
            Name="${internalName}" ${minString} ${maxString}>
            ${defaultString}
            </Field>`
        };

        props.saveButtonHandler(body);
    }

    const optionsDisplayFormat: IDropdownOption[] = [
        { key: -1, text: 'Automatic' },
        { key: 0, text: '0' },
        { key: 1, text: '1' },
        { key: 2, text: '2' },
        { key: 3, text: '3' },
        { key: 4, text: '4' },
        { key: 5, text: '5' }
      ];

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
                <TextField label="Minimum allowed value" type="number" onChange={(evt: React.FormEvent<HTMLInputElement>) => { 
                                setMinValue(((evt.target as any).valueAsNumber.toString().length == 0) ? null : (evt.target as any).valueAsNumber);}
                                } />
                            <TextField label="Maximum allowed value" type="number" onChange={(evt: React.FormEvent<HTMLInputElement>) => { 
                                setMaxValue(((evt.target as any).valueAsNumber.toString().length == 0) ? null : (evt.target as any).valueAsNumber);}
                                } />
                            <Dropdown label="Number of decimal places" options={optionsDisplayFormat} defaultSelectedKey={displayFormat} onChanged={(evt: IDropdownOption) => {
                                setDisplayFormat(+(evt.key));}
                            }/>                        
                            <TextField label="Default value" type="number" onChange={(evt: React.FormEvent<HTMLInputElement>) => { 
                                setDefaultValue((evt.target as any).valueAsNumber.toString());}
                                } />
                            <Toggle label="Show as percentage (for example, 50%)" onChanged={(evt) => setShowAsPercentage(evt)} />
                <br />
                <PrimaryButton text="Save" onClick={() => saveNewField()} />
                <Button text="Cancel" onClick={() => props.cancelButtonHandler()} />
            </>
    );
}