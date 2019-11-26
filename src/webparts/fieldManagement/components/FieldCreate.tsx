import * as React from 'react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { PrimaryButton, Button, Dropdown, IDropdownOption, FacepileBase } from 'office-ui-fabric-react';
import { ISPHttpClientOptions, SPHttpClientResponse, SPHttpClient } from '@microsoft/sp-http';
import { ISPField } from './SPField';
import { FieldTypeKindEnum } from './FieldTypeKindEnum';
import { BaseComponentContext } from '@microsoft/sp-component-base';

export interface FieldCreateProps {
    group: string,
    context: BaseComponentContext
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
    displayFormat: string
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
            displayFormat: "-1"
        }
    }



    generateInternalName(){
        let columnName: string = (document.getElementById("columnName") as HTMLInputElement).value;
        let newValue: string =  columnName.replace(/[^A-Z0-9]+/ig, "");
        this.setState({columnName: columnName, internalName: (newValue.length >= 32 ? newValue.substr(0, 32) : newValue)});
        //Add column internalName validation check. + possible counter
    }

    createFieldHandler(): Promise<any>{
        let context = this.props.context;
        let data = this.state;
        let body: ISPField;
        switch(this.state.fieldType){
            case FieldTypeKindEnum.Text:
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
                    Group: data.group
                }
            case FieldTypeKindEnum.Note:
                    body = {
                        "@odata.type": "#SP.FieldMultiLineText",
                        Title: data.columnName,
                        StaticName: data.internalName,
                        InternalName: data.internalName,
                        FieldTypeKind: FieldTypeKindEnum.Note,
                        Required: data.required,
                        Group: data.group,

                        UnlimitedLengthInDocumentLibrary: data.allowUnlimitedLength,
                        AllowHyperlink: data.allowRichText,
                        AppendOnly: data.appendChangesToExistingText,
                        NumberOfLines: data.numberOfLinesForEditing,
                        RichText: data.allowRichText
                    }
            case FieldTypeKindEnum.Number:
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
                    DisplayFormat: "2",
                    MaximumValue: data.maxValue == null ? 1.7976931348623157e+308 : data.maxValue,
                    MinimumValue: data.minValue == null ? -1.7976931348623157e+308 : data.minValue,
                    ShowAsPercentage: data.showAsPercentage,
                }
        }
        
        let bodyStr = JSON.stringify(body);
        console.log(bodyStr);
        const h2 = new Headers();
        h2.append("Accept", "application/json;odata.metadata=full");
        h2.append("Content-type", "application/json;odata.metadata=full");
    
        const optUpdate1: ISPHttpClientOptions = {
            headers: h2,
            body: bodyStr
        };
        return context.spHttpClient.post(context.pageContext.web.absoluteUrl + `/_api/web/fields`, SPHttpClient.configurations.v1, optUpdate1)
            .then((response: SPHttpClientResponse) => {
            console.log(response.json());
            return response.json();
        });
    }

    render() {
        const options: IDropdownOption[] = [
            { key: FieldTypeKindEnum.Text, text: 'Single line of text' },
            { key: FieldTypeKindEnum.Note, text: 'Multiple lines of text' },
            { key: FieldTypeKindEnum.Number, text: 'Number (1, 1.0, 100)' },
            { key: FieldTypeKindEnum.Choice , text: 'Choice (menu to choose from)', disabled: true },
            { key: FieldTypeKindEnum.Currency , text: 'Currency ($, ¥, €)', disabled: true },
            { key: FieldTypeKindEnum.DateTime , text: 'Date and Time', disabled: true },
            { key: FieldTypeKindEnum.Lookup , text: 'Lookup (information already on this site)', disabled: true },
            { key: FieldTypeKindEnum.Boolean , text: 'Yes/No (check box)', disabled: true },
            { key: FieldTypeKindEnum.User , text: 'Person or Group', disabled: true },
            { key: FieldTypeKindEnum.URL , text: 'Hyperlink or Picture', disabled: true },
            { key: FieldTypeKindEnum.Calculated , text: 'Calculated (calculation based on other columns)', disabled: true }
          ];
          const optionsDisplayFormat: IDropdownOption[] = [
            { key: '-1', text: 'Automatic' },
            { key: '0', text: '0' },
            { key: '1', text: '1' },
            { key: '2', text: '2' },
            { key: '3', text: '3' },
            { key: '4', text: '4' },
            { key: '5', text: '5' }
          ];
        return (
            <div>
                { 
                    this.state.fieldType == FieldTypeKindEnum.Text ?
                        <div>
                        <TextField label="Column Name" id="columnName" required value={this.state.columnName} onKeyUp={() => this.generateInternalName()} />
                        <Dropdown label="Field Type" options={options} defaultSelectedKey={this.state.fieldType} onChanged={(evt: any) => this.setState({fieldType: evt.key})} />
                        <TextField label="Internal Name" required value={this.state.internalName} onKeyUp={(evt) => this.setState({internalName: (evt.target as HTMLInputElement).value})} />
                        <TextField label="Group" defaultValue={this.props.group} onChanged={(evt: string) => { this.setState({ group: evt })}} />
                        <TextField label="Description" name="columnName" multiline autoAdjustHeight onChanged={(evt: string) => { this.setState({ description: evt })}} />
                        <Toggle label="Required" onChanged={(evt) => this.setState({required: evt})} />
                        <Toggle label="Enforce Unique Values" onChanged={(evt) => this.setState({enforceUniqueValues: evt})} />
                        <TextField label="Maximum number of characters" max={255} min={0} type="number" defaultValue="255" onChanged={(evt: number) => { this.setState({ maxNoCharacters: evt })}} />
                        <TextField label="Default value" value={this.state.defaultValue} onChanged={(evt: string) => { this.setState({ defaultValue: evt })}} />
                    </div> : null
                }
                {
                    this.state.fieldType == FieldTypeKindEnum.Note ?
                        <div>
                            <TextField label="Column Name" id="columnName" required value={this.state.columnName} onKeyUp={() => this.generateInternalName()} />
                            <Dropdown label="Field Type" options={options} defaultSelectedKey={this.state.fieldType} onChanged={(evt: any) => this.setState({fieldType: evt.key})} />
                            <TextField label="Internal Name" required value={this.state.internalName} onKeyUp={(evt) => this.setState({internalName: (evt.target as HTMLInputElement).value})} />
                            <TextField label="Group" defaultValue={this.props.group} onChanged={(evt: string) => { this.setState({ group: evt })}} />
                            <TextField label="Description" name="columnName" multiline autoAdjustHeight onChanged={(evt: string) => { this.setState({ description: evt })}} />
                            <Toggle label="Required" onChanged={(evt) => this.setState({required: evt})} />
                            <Toggle label="Allow unlimited length in document libraries" onChanged={(evt) => this.setState({allowUnlimitedLength: evt})} />
                            <TextField label="Number of lines for editing" max={255} min={0} type="number" defaultValue="6" onChanged={(evt: number) => { this.setState({ numberOfLinesForEditing: evt })}} />
                            <Toggle label="Allow enhanced rich text" checked={this.state.allowRichText} onChanged={(evt) => this.setState({allowRichText: evt})} /> 
                            <Toggle label="Append Changes to Existing Text" onChanged={(evt) => this.setState({appendChangesToExistingText: evt})} />
                        </div> : null
                }
                {
                    this.state.fieldType == FieldTypeKindEnum.Number ?
                        <div>
                            <TextField label="Column Name" id="columnName" required value={this.state.columnName} onKeyUp={() => this.generateInternalName()} />
                            <Dropdown label="Field Type" options={options} defaultSelectedKey={this.state.fieldType} onChanged={(evt: any) => this.setState({fieldType: evt.key})} />
                            <TextField label="Internal Name" required value={this.state.internalName} onKeyUp={(evt) => this.setState({internalName: (evt.target as HTMLInputElement).value})} />
                            <TextField label="Group" defaultValue={this.props.group} onChanged={(evt: string) => { this.setState({ group: evt })}} />
                            <TextField label="Description" name="columnName" multiline autoAdjustHeight onChanged={(evt: string) => { this.setState({ description: evt })}} />
                            <Toggle label="Required" onChanged={(evt) => this.setState({required: evt})} />
                            <Toggle label="Enforce Unique Values" onChanged={(evt) => this.setState({enforceUniqueValues: evt})} />
                            <TextField label="Minimum allowed value" type="number" onChanged={(evt: number) => { this.setState({ minValue: evt })}} />
                            <TextField label="Maximum allowed value" type="number" onChanged={(evt: number) => { this.setState({ maxValue: evt })}} />
                            <Dropdown label="Number of decimal places" options={optionsDisplayFormat} defaultSelectedKey={this.state.displayFormat} onChanged={(evt: any) => {
                                console.log({evt});
                                this.setState({displayFormat: evt.key})}
                                
                            }/>                        
                            <TextField label="Default value" type="number" value={this.state.defaultValue} onChanged={(evt: number) => { this.setState({ defaultValue: evt.toString() })}} />
                            <Toggle label="Show as percentage (for example, 50%)" onChanged={(evt) => this.setState({showAsPercentage: evt})} />
                        </div> : null
                }
            <br /><PrimaryButton text="Save" onClick={() => this.createFieldHandler()} />
                <Button text="Cancel" />
            </div>
        );
    }
}