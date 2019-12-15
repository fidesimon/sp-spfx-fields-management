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
import { CreateURLField } from './CreateFieldComponents/CreateURLField';
import { CreateDateTimeField } from './CreateFieldComponents/CreateDateTimeField';
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

    protected changeFieldType(fieldType: FieldTypeKindEnum){
        this.setState({fieldType: fieldType});
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
                    <CreateTextField saveButtonHandler={this.createNewField.bind(this)} groupName={this.props.group} fieldTypeOptions={options} cancelButtonHandler={this.props.closePanel} onFieldTypeChange={this.changeFieldType.bind(this)} />
                    : null
                }
                {
                    this.state.fieldType == FieldTypeKindEnum.Note ? 
                    <CreateMultiLineField saveButtonHandler={this.createNewField.bind(this)} groupName={this.props.group} fieldTypeOptions={options} cancelButtonHandler={this.props.closePanel} onFieldTypeChange={this.changeFieldType.bind(this)} />
                    : null
                }
                {
                    this.state.fieldType == FieldTypeKindEnum.Number ? 
                    <CreateNumberField saveButtonHandler={this.createNewField.bind(this)} groupName={this.props.group} fieldTypeOptions={options} cancelButtonHandler={this.props.closePanel} onFieldTypeChange={this.changeFieldType.bind(this)} />
                    : null
                }
                {
                    this.state.fieldType == FieldTypeKindEnum.Currency ? 
                    <CreateCurrencyField saveButtonHandler={this.createNewField.bind(this)} groupName={this.props.group} fieldTypeOptions={options} cancelButtonHandler={this.props.closePanel} onFieldTypeChange={this.changeFieldType.bind(this)} />
                    : null
                }
                {
                    this.state.fieldType == FieldTypeKindEnum.Choice ? 
                    <CreateChoiceField saveButtonHandler={this.createNewField.bind(this)} groupName={this.props.group} fieldTypeOptions={options} cancelButtonHandler={this.props.closePanel} onFieldTypeChange={this.changeFieldType.bind(this)} />
                    : null
                }
                {
                    this.state.fieldType == FieldTypeKindEnum.Boolean ? 
                    <CreateBooleanField saveButtonHandler={this.createNewField.bind(this)} groupName={this.props.group} fieldTypeOptions={options} cancelButtonHandler={this.props.closePanel} onFieldTypeChange={this.changeFieldType.bind(this)} />
                    : null
                }
                {
                    this.state.fieldType == FieldTypeKindEnum.URL ? 
                    <CreateURLField saveButtonHandler={this.createNewField.bind(this)} groupName={this.props.group} fieldTypeOptions={options} cancelButtonHandler={this.props.closePanel} onFieldTypeChange={this.changeFieldType.bind(this)} />
                    : null
                }
                {
                    this.state.fieldType == FieldTypeKindEnum.DateTime ? 
                    <CreateDateTimeField saveButtonHandler={this.createNewField.bind(this)} groupName={this.props.group} fieldTypeOptions={options} cancelButtonHandler={this.props.closePanel} onFieldTypeChange={this.changeFieldType.bind(this)} />
                    : null
                }
            </>
        );
    }
}