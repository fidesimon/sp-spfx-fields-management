import * as React from 'react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { PrimaryButton, Button } from 'office-ui-fabric-react';
import { ISPHttpClientOptions, SPHttpClientResponse, SPHttpClient } from '@microsoft/sp-http';
import { ISPField } from './SPField';
import { FieldTypeKindEnum } from './FieldTypeKindEnum';
import { BaseComponentContext } from '@microsoft/sp-component-base';

export interface FieldCreateProps {
    group: string,
    context: BaseComponentContext
}

export interface FieldCreateState{
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
        this.state = { columnName: '', internalName: '', group: this.props.group, description: '', required: false, enforceUniqueValues: false, maxNoCharacters: 255, defaultValue: ''}
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
        return (
            <div>
                <TextField label="Column Name" id="columnName" required value={this.state.columnName} onKeyUp={() => this.generateInternalName()} />
                <TextField label="Internal Name" required value={this.state.internalName} onKeyUp={(evt) => this.setState({internalName: (evt.target as HTMLInputElement).value})} />
                <TextField label="Group" defaultValue={this.props.group} onChanged={(evt: string) => { this.setState({ group: evt })}} />
                <TextField label="Description" name="columnName" multiline autoAdjustHeight onChanged={(evt: string) => { this.setState({ description: evt })}} />
                <Toggle label="Required" onChanged={(evt) => this.setState({required: evt})} />
                <Toggle label="Enforce Unique Values?" onChanged={(evt) => this.setState({enforceUniqueValues: evt})} />
                <TextField label="Maximum number of characters" max={255} min={0} type="number" defaultValue="255" onChanged={(evt: number) => { this.setState({ maxNoCharacters: evt })}} />
                <TextField label="Default value" value={this.state.defaultValue} onChanged={(evt: string) => { this.setState({ defaultValue: evt })}} />
                <br /><PrimaryButton text="Save" onClick={() => this.createFieldHandler()} />
                <Button text="Cancel" />
            </div>
        );
    }
}

/*
For TextField options in SharePoint:
1. Column Name
2. Group - existing group drop-down or new group
3. Description
4. Required
5. Enforce unique values
6. Maximum number of characters (def 255) 
7. Default Value - text or calculated value - calculated value skipped until v2
8. Column formatting (json) - skipped until v2
9. Column validation: Formula and User Message.   - skipped until v2

FieldTypeKind: 2


*/