import * as React from 'react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { PrimaryButton, Button, Dropdown, IDropdownOption, FacepileBase, IChoiceGroupOption, ChoiceGroup, IDropdown, DatePicker, DayOfWeek } from 'office-ui-fabric-react';
import { ISPHttpClientOptions, SPHttpClientResponse, SPHttpClient } from '@microsoft/sp-http';
import { CreateTextField } from './CreateFieldComponents/CreateTextField';
import { CreateMultiLineField } from './CreateFieldComponents/CreateMultiLineField';
import { CreateNumberField } from './CreateFieldComponents/CreateNumberField';
import { CreateCurrencyField } from './CreateFieldComponents/CreateCurrencyField';
import { CreateChoiceField } from './CreateFieldComponents/CreateChoiceField';
import { CreateBooleanField } from './CreateFieldComponents/CreateBooleanField';
import { ISPField } from './SPField';
import { FieldTypeKindEnum } from './FieldTypeKindEnum';
import { BaseComponentContext } from '@microsoft/sp-component-base';

export interface FieldCreateProps {
    group: string;
    context: BaseComponentContext;
    onItemSaved: Function;
    closePanel: Function;
}

export interface FieldCreateState{
    fieldType: number;
    columnName: string;
    internalName: string;
    group: string;
    description: string;
    required: boolean;
    enforceUniqueValues: boolean;
    maxNoCharacters: number;
    defaultValue: string;
    numberOfLinesForEditing: number;
    allowUnlimitedLength: boolean;
    allowRichText: boolean;
    appendChangesToExistingText: boolean;
    minValue: number;
    maxValue: number;
    showAsPercentage: boolean;
    displayFormat: number;
    choices: string[];
    choiceFormat: string;
    choiceFillIn: boolean;
    defaultValueChoices: IDropdownOption[];
    selectedCurrency: string;
    defaultBooleanValueAsString: string;
    urlFieldFormat: string;
    dateAndTimeFormat: string;
    friendlyDisplayFormat: string;
    displayDefaultValueDTInput: boolean;
}

export default class FieldCreate extends React.Component<FieldCreateProps, FieldCreateState>{
    constructor(props){
        super(props);
        this.state = { fieldType: FieldTypeKindEnum.Text, 
            columnName: '', 
            internalName: '', 
            group: this.props.group, 
            description: '', 
            required: false, 
            enforceUniqueValues: false, 
            maxNoCharacters: 255, 
            defaultValue: '',
            numberOfLinesForEditing: 6,
            allowUnlimitedLength: false,
            allowRichText: true,
            appendChangesToExistingText: false,
            minValue: null,
            maxValue: null,
            showAsPercentage: false,
            displayFormat: -1,
            choices: ["Enter Choice #1", "Enter Choice #2", "Enter Choice #3"],
            choiceFormat: 'Dropdown',
            choiceFillIn: false,
            selectedCurrency: "1033",
            defaultBooleanValueAsString: "1",
            urlFieldFormat: "Hyperlink",
            dateAndTimeFormat: "DateOnly",
            friendlyDisplayFormat: "Disabled",
            displayDefaultValueDTInput: false,
            defaultValueChoices: [{key: '', text: '(empty)', isSelected: true}, {key: 'Enter Choice #1', text: 'Enter Choice #1'}, {key: 'Enter Choice #2', text: 'Enter Choice #2'},{key: 'Enter Choice #3', text: 'Enter Choice #3'}]
        };
    }

    protected generateInternalName(){
        let columnName: string = (document.getElementById("columnName") as HTMLInputElement).value;
        let newValue: string =  columnName.replace(/[^A-Z0-9]+/ig, "");
        this.setState({columnName: columnName, internalName: (newValue.length >= 32 ? newValue.substr(0, 32) : newValue)});
        //Add column internalName validation check. + possible counter
    }

    private getUpperCaseStringForBool = (value: boolean) => value.toString().toUpperCase();

    protected async createFieldHandler(): Promise<any>{
        let data = this.state;
        let body: ISPField;
        let defaultString: string = '';
        switch(this.state.fieldType){
            case FieldTypeKindEnum.URL:
                    body = {
                        "@odata.type": "#SP.FieldUrl",
                        Title: data.columnName,
                        StaticName: data.internalName,
                        InternalName: data.internalName,
                        FieldTypeKind: FieldTypeKindEnum.URL,
                        Required: data.required,
                        Group: data.group,
                        Description: data.description,
                        SchemaXml: '<Field Type="URL" DisplayName="'+ data.columnName + '" Format="'+data.urlFieldFormat+'" Required="'+this.getUpperCaseStringForBool(data.required)+'" EnforceUniqueValues="FALSE" Description="'+data.description+'" Group="'+data.group+'" StaticName="'+data.internalName+'" Name="'+data.internalName+'"></Field>'
                    };
            break;
            case FieldTypeKindEnum.DateTime:
                    defaultString = data.defaultValue == "" ? "" : `<Default>${data.defaultValue}</Default>`;
                    let additionalAttributesForToday = data.defaultValue == "[today]" ? `CustomFormatter="" CalType="0"` : ``;
                    body = {
                        "@odata.type": "#SP.FieldDateTime",
                        Title: data.columnName,
                        StaticName: data.internalName,
                        InternalName: data.internalName,
                        FieldTypeKind: FieldTypeKindEnum.DateTime,
                        Required: data.required,
                        EnforceUniqueValues: data.enforceUniqueValues,
                        Group: data.group,
                        Description: data.description,
                        SchemaXml: `<Field Type="DateTime" DisplayName="${data.columnName}" Required="${this.getUpperCaseStringForBool(data.required)}" ${additionalAttributesForToday} EnforceUniqueValues="${this.getUpperCaseStringForBool(data.enforceUniqueValues)}" Format="${data.dateAndTimeFormat}" Group="${data.group}" FriendlyDisplayFormat="${data.friendlyDisplayFormat}" StaticName="${data.internalName}" Name="${data.internalName}">${defaultString}</Field>`
                    };
            break;
        }
        
        await this.createNewField(body);
    }

    protected async createNewField(body: ISPField): Promise<any>{
        let context = this.props.context;

        let bodyStr = JSON.stringify(body);
        const headers = new Headers();
        headers.append("Accept", "application/json;odata.metadata=full");
        headers.append("Content-type", "application/json;odata.metadata=full");
    
        const optUpdate1: ISPHttpClientOptions = {
            headers: headers,
            body: bodyStr
        };
        let response = await context.spHttpClient.post(context.pageContext.web.absoluteUrl + `/_api/web/fields`, SPHttpClient.configurations.v1, optUpdate1);
        let jsonResponse = await response.json();
        if(response.status == 201){
            this.props.onItemSaved(jsonResponse);
        }
    }

    // Comment for rendering part: There are components with the following logic:
    // (evt.toString().length == 0) ? null : evt
    // where evt is number. Through the code when entered value is removed (backspace)
    // the evt value would be assigned evt = "", but because it's not a number it cannot be checked otherwise
    // whether the value is (null or empty) or contains a value. hence the converting.
    render() {
        const options: IDropdownOption[] = [
            { key: FieldTypeKindEnum.Text, text: 'Single line of text' },
            { key: FieldTypeKindEnum.Note, text: 'Multiple lines of text' },
            { key: FieldTypeKindEnum.Number, text: 'Number (1, 1.0, 100)' },
            { key: FieldTypeKindEnum.Choice , text: 'Choice (menu to choose from)' },
            { key: FieldTypeKindEnum.Currency , text: 'Currency ($, ¥, €)' },
            { key: FieldTypeKindEnum.DateTime , text: 'Date and Time' },
            { key: FieldTypeKindEnum.Lookup , text: 'Lookup (information already on this site)', disabled: true },
            { key: FieldTypeKindEnum.Boolean , text: 'Yes/No (check box)' },
            { key: FieldTypeKindEnum.User , text: 'Person or Group', disabled: true },
            { key: FieldTypeKindEnum.URL , text: 'Hyperlink or Picture' },
            { key: FieldTypeKindEnum.Calculated , text: 'Calculated (calculation based on other columns)', disabled: true }
          ];

        return (
            <>
                {
                    this.state.fieldType == FieldTypeKindEnum.Text ? 
                    <CreateTextField saveButtonHandler={this.createNewField.bind(this)} groupName={this.props.group} fieldTypeOptions={options} cancelButtonHandler={this.props.closePanel} />
                    : null
                }
                {
                    this.state.fieldType == FieldTypeKindEnum.Note ? 
                    <CreateMultiLineField saveButtonHandler={this.createNewField.bind(this)} groupName={this.props.group} fieldTypeOptions={options} cancelButtonHandler={this.props.closePanel} />
                    : null
                }
                {
                    this.state.fieldType == FieldTypeKindEnum.Number ? 
                    <CreateNumberField saveButtonHandler={this.createNewField.bind(this)} groupName={this.props.group} fieldTypeOptions={options} cancelButtonHandler={this.props.closePanel} />
                    : null
                }
                {
                    this.state.fieldType == FieldTypeKindEnum.Currency ? 
                    <CreateCurrencyField saveButtonHandler={this.createNewField.bind(this)} groupName={this.props.group} fieldTypeOptions={options} cancelButtonHandler={this.props.closePanel} />
                    : null
                }
                {
                    this.state.fieldType == FieldTypeKindEnum.Choice ? 
                    <CreateChoiceField saveButtonHandler={this.createNewField.bind(this)} groupName={this.props.group} fieldTypeOptions={options} cancelButtonHandler={this.props.closePanel} />
                    : null
                }
                {
                    this.state.fieldType == FieldTypeKindEnum.Boolean ? 
                    <CreateBooleanField saveButtonHandler={this.createNewField.bind(this)} groupName={this.props.group} fieldTypeOptions={options} cancelButtonHandler={this.props.closePanel} />
                    : null
                }
                <TextField label="Column Name" id="columnName" required value={this.state.columnName} onKeyUp={() => this.generateInternalName()} />
                <Dropdown label="Field Type" options={options} defaultSelectedKey={this.state.fieldType} onChanged={(evt: any) => this.setState({fieldType: evt.key})} />
                <TextField label="Internal Name" required value={this.state.internalName} onKeyUp={(evt) => this.setState({internalName: (evt.target as HTMLInputElement).value})} />
                <TextField label="Group" defaultValue={this.props.group} onChange={(evt: React.FormEvent<HTMLInputElement>) => { this.setState({ group: (evt.target as any).value });}} />
                <TextField label="Description" name="columnName" multiline autoAdjustHeight onChange={(evt: React.FormEvent<HTMLTextAreaElement>) => { this.setState({ description: (evt.target as any).value });}} />
                {
                    this.state.fieldType == FieldTypeKindEnum.URL ? 
                    <>
                        <Toggle label="Required" onChanged={(evt) => this.setState({required: evt})} />
                        <Dropdown label="Format URL as" options={[{key: 'Hyperlink', text: 'Hyperlink'}, {key: 'Image', text: 'Picture'}]} defaultSelectedKey={this.state.urlFieldFormat} onChanged={(evt:any) => {
                            this.setState({urlFieldFormat: evt.key});
                        }} />
                    </>
                    : null
                }
                {
                    this.state.fieldType == FieldTypeKindEnum.DateTime ?
                    <>
                        <Toggle label="Required" onChanged={(evt) => this.setState({required: evt})} />
                        <Toggle label="Enforce Unique Values" onChanged={(evt) => this.setState({enforceUniqueValues: evt})} />
                        <ChoiceGroup styles={{flexContainer: {display: "flex"}}} label="Date and Time Format" defaultSelectedKey={this.state.dateAndTimeFormat} options={[{key: "DateOnly", text: "Date Only\u00A0\u00A0"},{key: "DateTime", text: "Date & Time"}]} onChanged={(evt: any) => { 
                                this.setState({dateAndTimeFormat: evt.key});
                            }} />
                        <ChoiceGroup styles={{flexContainer: {display: "flex"}}} label="Display Format" defaultSelectedKey={this.state.friendlyDisplayFormat} options={[{key: "Disabled", text: "Standard\u00A0\u00A0"},{key: "Relative", text: "Friendly"}]} onChanged={(evt: any) => { 
                                this.setState({friendlyDisplayFormat: evt.key});
                            }} />
                        <ChoiceGroup label="Default Value" defaultSelectedKey="None" options={[{key: "None", text: "(None)"},{key: "[today]", text: "Today's Date"}, {key: "Another", text: "Specified Date"}]} onChanged={(evt: any) => { 
                                switch(evt.key){
                                    case "None":
                                        this.setState({defaultValue: "", displayDefaultValueDTInput: false});
                                    break;
                                    case "[today]":
                                        this.setState({defaultValue: "[today]", displayDefaultValueDTInput: false});
                                    break;
                                    case "Another":
                                        this.setState({defaultValue: "", displayDefaultValueDTInput: true});
                                    break;
                                }
                                this.setState({dateAndTimeFormat: evt.key});
                            }} />
                            {this.state.displayDefaultValueDTInput ? 
                                <DatePicker 
                                label="Enter Default Date" 
                                defaultValue={this.state.selectedCurrency}
                                allowTextInput={false} 
                                firstDayOfWeek={DayOfWeek.Monday} 
                                formatDate={this._onFormatDate}
                                onSelectDate={this._onSelectDate}
                                value={this.state.defaultValue == "" ? null : new Date(this.state.defaultValue)}
                                />
                                : null
                            }
                    </>
                    : null
                }
            <br /><PrimaryButton text="Save" onClick={() => this.createFieldHandler()} />
                <Button text="Cancel" onClick={() => this.props.closePanel()} />
            </>
        );
    }
    private _onFormatDate = (date: Date): string => {
        return date.getDate() + '/' + (date.getMonth() + 1) + '/' + (date.getFullYear() % 100);
    }

    private _onSelectDate = (date: Date | null | undefined): void => {
        //Need to compensate the time difference between GMT and LocaleTime
        this.setState({defaultValue: (new Date(date.getTime() + Math.abs(date.getTimezoneOffset() * (-60000))).toISOString())});
    }
}