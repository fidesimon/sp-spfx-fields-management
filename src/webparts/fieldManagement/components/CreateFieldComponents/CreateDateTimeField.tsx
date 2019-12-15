import * as React from 'react';
import { PrimaryButton, Button, Dropdown, IDropdownOption, TextField, Toggle, ChoiceGroup, DatePicker, DayOfWeek, IChoiceGroupOption } from 'office-ui-fabric-react';
import { FieldTypeKindEnum } from '../FieldTypeKindEnum';
import { ISPField } from '../SPField';

export interface CreateDateTimeFieldProps {
    fieldTypeOptions: IDropdownOption[];
    saveButtonHandler: Function;
    cancelButtonHandler: Function;
    groupName: string;
    onFieldTypeChange: Function;
}

export const CreateDateTimeField: React.FC<CreateDateTimeFieldProps> = (props) => {
    const [columnName, setColumnName] = React.useState("");
    const [fieldType, setFieldType] = React.useState(FieldTypeKindEnum.DateTime);
    const [internalName, setInternalName] = React.useState("");
    const [group, setGroup] = React.useState(props.groupName);
    const [description, setDescription] = React.useState("");
    const [required,setRequired] = React.useState(false);
    const [enforceUniqueValues, setEnforceUniqueValues] = React.useState<boolean>(false);
    const [defaultValue, setDefaultValue] = React.useState();
    const [dateAndTimeFormat, setDateAndTimeFormat] = React.useState("DateOnly");
    const [friendlyDisplayFormat, setFriendlyDisplayFormat] = React.useState("Disabled");
    const [displayDefaultValueDTInput, setDisplayDefaultValueDTInput] = React.useState(false);

    const saveNewField = () => {
        let body: ISPField;
        let defaultString: string = (defaultValue == "") ? "": `<Default>${defaultValue}</Default>`;
        let additionalAttributesForToday = defaultValue == "[today]" ? `CustomFormatter="" CalType="0"` : ``;

        body = {
            "@odata.type": "#SP.FieldDateTime",
            Title: columnName,
            StaticName: internalName,
            InternalName: internalName,
            FieldTypeKind: FieldTypeKindEnum.DateTime,
            Required: required,
            EnforceUniqueValues: enforceUniqueValues,
            DefaultValue: defaultValue,
            Group: group,
            Description: description,
            SchemaXml: `<Field Type="DateTime" Description="${description}" DisplayName="${columnName}" Required="${(required ? "TRUE" : "FALSE")}" ${additionalAttributesForToday} EnforceUniqueValues="${(enforceUniqueValues? "TRUE" : "FALSE")}" Group="${group}" StaticName="${internalName}" Name="${internalName}" Format="${dateAndTimeFormat}" FriendlyDisplayFormat="${friendlyDisplayFormat}">${defaultString}</Field>`
        };

        props.saveButtonHandler(body);
    }

    const dateAndTimeFormatChoices: IChoiceGroupOption[] = [
        {key: "DateOnly", text: "Date Only\u00A0\u00A0"},
        {key: "DateTime", text: "Date & Time"}
    ];

    const friendlyDisplayFormatChoices: IChoiceGroupOption[] = [
        {key: "Disabled", text: "Standard\u00A0\u00A0"},
        {key: "Relative", text: "Friendly"}
    ];

    const defaultValueOptions: IChoiceGroupOption[] = [
        {key: "None", text: "(None)"},
        {key: "[today]", text: "Today's Date"}, 
        {key: "Another", text: "Specified Date"}
    ];

    const onFormatDate = (date: Date): string => {
        return date.getDate() + '/' + (date.getMonth() + 1) + '/' + (date.getFullYear() % 100);
    }

    const onSelectDate = (date: Date | null | undefined): void => {
        //Need to compensate the time difference between GMT and LocaleTime
        setDefaultValue((new Date(date.getTime() + Math.abs(date.getTimezoneOffset() * (-60000))).toISOString()));
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
                <Toggle label="Enforce Unique Values" onChanged={(evt) => setEnforceUniqueValues(evt)} />
                <ChoiceGroup styles={{flexContainer: {display: "flex"}}} label="Date and Time Format" defaultSelectedKey={dateAndTimeFormat} options={dateAndTimeFormatChoices} onChanged={(evt: any) => { 
                        setDateAndTimeFormat(evt.key);
                    }} />
                <ChoiceGroup styles={{flexContainer: {display: "flex"}}} label="Display Format" defaultSelectedKey={friendlyDisplayFormat} options={friendlyDisplayFormatChoices} onChanged={(evt: any) => { 
                        setFriendlyDisplayFormat(evt.key);
                    }} />
                <ChoiceGroup label="Default Value" defaultSelectedKey="None" options={defaultValueOptions} onChanged={(evt: any) => { 
                        switch(evt.key){
                            case "None":
                                setDefaultValue("");
                                setDisplayDefaultValueDTInput(false);
                            break;
                            case "[today]":
                                setDefaultValue("[today]");
                                setDisplayDefaultValueDTInput(false);
                            break;
                            case "Another":
                                setDefaultValue("");
                                setDisplayDefaultValueDTInput(true);
                            break;
                        }
                        this.setState({dateAndTimeFormat: evt.key});
                    }} />
                    {displayDefaultValueDTInput ? 
                        <DatePicker 
                        label="Enter Default Date" 
                        allowTextInput={false} 
                        firstDayOfWeek={DayOfWeek.Monday} 
                        formatDate={onFormatDate}
                        onSelectDate={onSelectDate}
                        value={defaultValue == "" ? null : new Date(defaultValue)}
                        />
                        : null
                    }
                <br />
                <PrimaryButton text="Save" onClick={() => saveNewField()} />
                <Button text="Cancel" onClick={() => props.cancelButtonHandler()} />
            </>
    );
}