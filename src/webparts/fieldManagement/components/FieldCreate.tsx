import * as React from 'react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { PrimaryButton, Button, Dropdown, IDropdownOption } from 'office-ui-fabric-react';
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
    defaultValue: string
}

export default class FieldCreate extends React.Component<FieldCreateProps, FieldCreateState>{
    constructor(props){
        super(props);
        this.state = { fieldType: FieldTypeKindEnum.Note, columnName: '', internalName: '', group: this.props.group, description: '', required: false, enforceUniqueValues: false, maxNoCharacters: 255, defaultValue: ''}
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
        let body: ISPField = {
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
            { key: FieldTypeKindEnum.Note, text: 'Multiple lines of text', disabled: true },
            { key: FieldTypeKindEnum.Number, text: 'Number (1, 1.0, 100)', disabled: true },
            { key: FieldTypeKindEnum.Choice , text: 'Choice (menu to choose from)', disabled: true },
            { key: FieldTypeKindEnum.Currency , text: 'Currency ($, ¥, €)', disabled: true },
            { key: FieldTypeKindEnum.DateTime , text: 'Date and Time', disabled: true },
            { key: FieldTypeKindEnum.Lookup , text: 'Lookup (information already on this site)', disabled: true },
            { key: FieldTypeKindEnum.Boolean , text: 'Yes/No (check box)', disabled: true },
            { key: FieldTypeKindEnum.User , text: 'Person or Group', disabled: true },
            { key: FieldTypeKindEnum.URL , text: 'Hyperlink or Picture', disabled: true },
            { key: FieldTypeKindEnum.Calculated , text: 'Calculated (calculation based on other columns)', disabled: true }
          ];
        return (
            <div>
                <TextField label="Column Name" id="columnName" required value={this.state.columnName} onKeyUp={() => this.generateInternalName()} />
                <Dropdown label="Field Type" options={options} defaultSelectedKey={this.state.fieldType} onChanged={(evt: any) => this.setState({fieldType: evt.key})} />
                <TextField label="Internal Name" required value={this.state.internalName} onKeyUp={(evt) => this.setState({internalName: (evt.target as HTMLInputElement).value})} />
                <TextField label="Group" defaultValue={this.props.group} onChanged={(evt: string) => { this.setState({ group: evt })}} />
                <TextField label="Description" name="columnName" multiline autoAdjustHeight onChanged={(evt: string) => { this.setState({ description: evt })}} />
                <Toggle label="Required" onChanged={(evt) => this.setState({required: evt})} />
                <Toggle label="Enforce Unique Values?" onChanged={(evt) => this.setState({enforceUniqueValues: evt})} />
                {this.state.fieldType == FieldTypeKindEnum.Text ? <TextField label="Maximum number of characters" max={255} min={0} type="number" defaultValue="255" onChanged={(evt: number) => { this.setState({ maxNoCharacters: evt })}} /> : null}
                <TextField label="Default value" value={this.state.defaultValue} onChanged={(evt: string) => { this.setState({ defaultValue: evt })}} />
                <br /><PrimaryButton text="Save" onClick={() => this.createFieldHandler()} />
                <Button text="Cancel" />
            </div>
        );
    }
}