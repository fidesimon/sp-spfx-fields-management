import * as React from 'react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { PrimaryButton, Button, Dropdown, IDropdownOption, FacepileBase, IChoiceGroupOption, ChoiceGroup } from 'office-ui-fabric-react';
import { ISPHttpClientOptions, SPHttpClientResponse, SPHttpClient } from '@microsoft/sp-http';
import { ISPField } from './SPField';
import { FieldTypeKindEnum } from './FieldTypeKindEnum';
import { BaseComponentContext } from '@microsoft/sp-component-base';

export interface FieldCreateProps {
    group: string,
    context: BaseComponentContext,
    onItemSaved: Function
}

export interface FieldCreateState{
    fieldType: number,
    columnName: string,
    internalName: string,
    group: string,
    description: string,
    required: boolean,
    enforceUniqueValues: boolean,
    maxNoCharacters: number,
    defaultValue: string,
    numberOfLinesForEditing: number,
    allowUnlimitedLength: boolean,
    allowRichText: boolean,
    appendChangesToExistingText: boolean,
    minValue: number,
    maxValue: number,
    showAsPercentage: boolean,
    displayFormat: number,
    choices: string[],
    choiceFormat: string,
    choiceFillIn: boolean,
    defaultValueChoices: IDropdownOption[]
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
            defaultValueChoices: [{key: '', text: '(empty)', isSelected: true}, {key: 'Enter Choice #1', text: 'Enter Choice #1'}, {key: 'Enter Choice #2', text: 'Enter Choice #2'},{key: 'Enter Choice #3', text: 'Enter Choice #3'}]
        }
    }

    generateInternalName(){
        let columnName: string = (document.getElementById("columnName") as HTMLInputElement).value;
        let newValue: string =  columnName.replace(/[^A-Z0-9]+/ig, "");
        this.setState({columnName: columnName, internalName: (newValue.length >= 32 ? newValue.substr(0, 32) : newValue)});
        //Add column internalName validation check. + possible counter
    }

    getUpperCaseStringForBool = (value: boolean) => value.toString().toUpperCase();

    async createFieldHandler(): Promise<any>{
        let context = this.props.context;
        let data = this.state;
        let body: ISPField;
        switch(this.state.fieldType){
            case FieldTypeKindEnum.Text:
                let defaultValueString = data.defaultValue.length == 0 ? '' : "<Default>" + data.defaultValue + "</Default>";
                body = {
                    "@odata.type": "#SP.FieldText",
                    Title: data.columnName,
                    StaticName: data.internalName,
                    InternalName: data.internalName,
                    FieldTypeKind: FieldTypeKindEnum.Text,
                    Required: data.required,
                    EnforceUniqueValues: data.enforceUniqueValues,
                    MaxLength: data.maxNoCharacters,
                    DefaultValue: data.defaultValue,
                    Group: data.group,
                    Description: data.description,
                    SchemaXml: '<Field Type="Text" Description="'+data.description+'" DisplayName="'+ data.columnName + '" Required="'+ (data.required? "TRUE" : "FALSE") +'" EnforceUniqueValues="'+ (data.enforceUniqueValues? "TRUE" : "FALSE") +'" Group="'+data.group+'" StaticName="'+data.internalName+'" Name="'+data.internalName+'">'+ defaultValueString +'</Field>'
                }
                break;
            case FieldTypeKindEnum.Note:
                body = {
                    "@odata.type": "#SP.FieldMultiLineText",
                    Title: data.columnName,
                    StaticName: data.internalName,
                    InternalName: data.internalName,
                    FieldTypeKind: FieldTypeKindEnum.Note,
                    Required: data.required,
                    Group: data.group,
                    Description: data.description,

                    UnlimitedLengthInDocumentLibrary: data.allowUnlimitedLength,
                    AppendOnly: data.appendChangesToExistingText,
                    NumberOfLines: data.numberOfLinesForEditing,
                    RichText: data.allowRichText,
                    SchemaXml: `<Field
                                Name="${data.internalName}"
                                DisplayName="${data.columnName}"
                                Description="${data.description}"
                                StaticName="${data.internalName}"
                                Group="${data.group}"
                                Type="Note"
                                NumLines="${data.numberOfLinesForEditing}"
                                UnlimitedLengthInDocumentLibrary="${this.getUpperCaseStringForBool(data.allowUnlimitedLength)}" 
                                Required="${this.getUpperCaseStringForBool(data.required)}" 
                                AppendOnly="${this.getUpperCaseStringForBool(data.appendChangesToExistingText)}" 
                                RichText="${this.getUpperCaseStringForBool(data.allowRichText)}"
                                />`
                }
                break;
            case FieldTypeKindEnum.Number:
                let minString = data.minValue == null ? '' : (data.showAsPercentage ? 'Min="' + data.minValue/100 + '"' : 'Min="' + data.minValue + '"');
                let maxString = data.maxValue == null ? '' : (data.showAsPercentage ? 'Max="' + data.maxValue/100 + '"' : 'Max="' + data.maxValue + '"');
                let defaultString = data.defaultValue.length == 0 ? '' : "<Default>" + (data.showAsPercentage ? (+(data.defaultValue)/100).toString() : data.defaultValue) + "</Default>";
                body = {
                    "@odata.type": "#SP.FieldNumber",
                    Title: data.columnName,
                    StaticName: data.internalName,
                    InternalName: data.internalName,
                    FieldTypeKind: FieldTypeKindEnum.Number,
                    Required: data.required,
                    EnforceUniqueValues: data.enforceUniqueValues,
                    DefaultValue: data.defaultValue,
                    Group: data.group,
                    DisplayFormat: +(data.displayFormat),
                    ShowAsPercentage: data.showAsPercentage,
                    Description: data.description,
                    SchemaXml: '<Field Type="Number" DisplayName="'+ data.columnName + '" Description="'+data.description+'" Required="'+ (data.required? "TRUE" : "FALSE") +'" Percentage="'+ (data.showAsPercentage? "TRUE" : "FALSE") +'" EnforceUniqueValues="'+ (data.enforceUniqueValues? "TRUE" : "FALSE") +'" Decimals="'+data.displayFormat+'" Group="'+data.group+'" StaticName="'+data.internalName+'" Name="'+data.internalName+'" Version="1" '+ minString + ' ' + maxString + '>'+ defaultString +'</Field>'
                }
                break;
            case FieldTypeKindEnum.Choice:
                let choicesString = `<CHOICES><CHOICE>${this.state.choices.join("</CHOICE><CHOICE>")}</CHOICE></CHOICES>`;
                let defaultChoiceValueString = (this.state.defaultValue == null || this.state.defaultValue == '') ? '' : `<Default>${this.state.defaultValue}</Default>`;
                body = {
                    "@odata.type": "#SP.FieldChoice",
                    Title: data.columnName,
                    StaticName: data.internalName,
                    InternalName: data.internalName,
                    FieldTypeKind: FieldTypeKindEnum.Choice,
                    Required: data.required,
                    Group: data.group,
                    DefaultValue: data.defaultValue,
                    EnforceUniqueValues: data.enforceUniqueValues,
                    Description: data.description,
                    SchemaXml: `<Field Type="Choice" DisplayName="${data.columnName}" StaticName="${data.internalName}" Description="${data.description}"  Name="${data.internalName}" Group="${data.group}" Format="${data.choiceFormat}" FillInChoice="${this.getUpperCaseStringForBool(data.choiceFillIn)}" Required="${this.getUpperCaseStringForBool(data.required)}" EnforceUniqueValues="${this.getUpperCaseStringForBool(data.enforceUniqueValues)}" >${defaultChoiceValueString}${choicesString}</Field>`
                }
                break;
        }
        
        let bodyStr = JSON.stringify(body);
        const h2 = new Headers();
        h2.append("Accept", "application/json;odata.metadata=full");
        h2.append("Content-type", "application/json;odata.metadata=full");
    
        const optUpdate1: ISPHttpClientOptions = {
            headers: h2,
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
            { key: FieldTypeKindEnum.Currency , text: 'Currency ($, ¥, €)', disabled: true },
            { key: FieldTypeKindEnum.DateTime , text: 'Date and Time', disabled: true },
            { key: FieldTypeKindEnum.Lookup , text: 'Lookup (information already on this site)', disabled: true },
            { key: FieldTypeKindEnum.Boolean , text: 'Yes/No (check box)', disabled: true },
            { key: FieldTypeKindEnum.User , text: 'Person or Group', disabled: true },
            { key: FieldTypeKindEnum.URL , text: 'Hyperlink or Picture', disabled: true },
            { key: FieldTypeKindEnum.Calculated , text: 'Calculated (calculation based on other columns)', disabled: true }
          ];
          const optionsDisplayFormat: IDropdownOption[] = [
            { key: -1, text: 'Automatic' },
            { key: 0, text: '0' },
            { key: 1, text: '1' },
            { key: 2, text: '2' },
            { key: 3, text: '3' },
            { key: 4, text: '4' },
            { key: 5, text: '5' }
          ];
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
          }
        return (
            <>
                <TextField label="Column Name" id="columnName" required value={this.state.columnName} onKeyUp={() => this.generateInternalName()} />
                <Dropdown label="Field Type" options={options} defaultSelectedKey={this.state.fieldType} onChanged={(evt: any) => this.setState({fieldType: evt.key})} />
                <TextField label="Internal Name" required value={this.state.internalName} onKeyUp={(evt) => this.setState({internalName: (evt.target as HTMLInputElement).value})} />
                <TextField label="Group" defaultValue={this.props.group} onChange={(evt: React.FormEvent<HTMLInputElement>) => { this.setState({ group: (evt.target as any).value })}} />
                <TextField label="Description" name="columnName" multiline autoAdjustHeight onChange={(evt: React.FormEvent<HTMLTextAreaElement>) => { this.setState({ description: (evt.target as any).value })}} />
                <Toggle label="Required" onChanged={(evt) => this.setState({required: evt})} />
                { 
                    this.state.fieldType == FieldTypeKindEnum.Text ?
                        <>
                            <Toggle label="Enforce Unique Values" onChanged={(evt) => this.setState({enforceUniqueValues: evt})} />
                            <TextField label="Maximum number of characters" max={255} min={0} type="number" defaultValue="255" onChange={(evt: React.FormEvent<HTMLInputElement>) => { this.setState({ maxNoCharacters: +((evt.target as any).value) })}} />
                            <TextField label="Default value" value={this.state.defaultValue} onChange={(evt: React.FormEvent<HTMLInputElement>) => { this.setState({ defaultValue: (evt.target as any).value })}} />
                        </> : null
                }
                {
                    this.state.fieldType == FieldTypeKindEnum.Note ?
                        <>
                            <Toggle label="Allow unlimited length in document libraries" onChanged={(evt) => this.setState({allowUnlimitedLength: evt})} />
                            <TextField label="Number of lines for editing" max={255} min={0} type="number" defaultValue="6" onChange={(evt: React.FormEvent<HTMLInputElement>) => { this.setState({ numberOfLinesForEditing: +((evt.target as any).value) })}} />
                            <Toggle label="Allow enhanced rich text" checked={this.state.allowRichText} onChanged={(evt) => {
                                this.setState({allowRichText: evt})}
                                } /> 
                            <Toggle label="Append Changes to Existing Text" onChanged={(evt) => this.setState({appendChangesToExistingText: evt})} />
                        </> : null
                }
                {
                    this.state.fieldType == FieldTypeKindEnum.Number ?
                        <>
                            <Toggle label="Enforce Unique Values" onChanged={(evt) => this.setState({enforceUniqueValues: evt})} />
                            <TextField label="Minimum allowed value" type="number" onChange={(evt: React.FormEvent<HTMLInputElement>) => { 
                                this.setState({ minValue: ((evt.target as any).valueAsNumber.toString().length == 0) ? null : (evt.target as any).valueAsNumber })}
                                } />
                            <TextField label="Maximum allowed value" type="number" onChange={(evt: React.FormEvent<HTMLInputElement>) => { 
                                this.setState({ maxValue: ((evt.target as any).valueAsNumber.toString().length == 0) ? null : (evt.target as any).valueAsNumber })}
                                } />
                            <Dropdown label="Number of decimal places" options={optionsDisplayFormat} defaultSelectedKey={this.state.displayFormat} onChanged={(evt: IDropdownOption) => {
                                this.setState({displayFormat: +(evt.key)})}
                            }/>                        
                            <TextField label="Default value" type="number" onChange={(evt: React.FormEvent<HTMLInputElement>) => { 
                                this.setState({ defaultValue: (evt.target as any).valueAsNumber.toString() })}
                                } />
                            <Toggle label="Show as percentage (for example, 50%)" onChanged={(evt) => this.setState({showAsPercentage: evt})} />
                        </> : null
                }
                {
                    this.state.fieldType == FieldTypeKindEnum.Choice ?
                        <>
                            <Toggle label="Enforce Unique Values" onChanged={(evt) => this.setState({enforceUniqueValues: evt})} />
                            <TextField 
                                label="Type each choice on a separate line" 
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
                                        this.setState({choices: distinctChoices, defaultValueChoices: defaultValueChoices});
                                    } 
                                }
                            />
                            <Dropdown label="Default value" defaultValue="(empty)" options={this.state.defaultValueChoices} onChanged={(evt: any) => {
                                this.setState({defaultValue: evt.key})
                            }} />
                            <ChoiceGroup label="Display choices using" defaultSelectedKey={this.state.choiceFormat} options={choiceFieldFormatOptions} onChanged={(evt: any) => { 
                                this.setState({choiceFormat: evt.key})
                            }} />
                            <Toggle label="Allow 'Fill-in' choices" onChanged={(evt) => this.setState({choiceFillIn: evt})} />
                        </> : null
                }
            <br /><PrimaryButton text="Save" onClick={() => this.createFieldHandler()} />
                <Button text="Cancel" />
            </>
        );
    }
}