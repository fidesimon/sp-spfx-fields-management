import * as React from 'react';
import { IDropdownOption } from 'office-ui-fabric-react';
import { ISPHttpClientOptions, SPHttpClient } from '@microsoft/sp-http';
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
}

export default class FieldCreate extends React.Component<FieldCreateProps, FieldCreateState>{
    constructor(props){
        super(props);
        this.state = { 
            fieldType: FieldTypeKindEnum.Text //default field type to start with
        };
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

        const renderCreateFieldComponent = (fieldType: FieldTypeKindEnum) => {
            switch(fieldType){
                case FieldTypeKindEnum.Text:
                    return <CreateTextField saveButtonHandler={this.createNewField.bind(this)} groupName={this.props.group} fieldTypeOptions={options} cancelButtonHandler={this.props.closePanel} onFieldTypeChange={this.changeFieldType.bind(this)} />
                case FieldTypeKindEnum.Note:
                    return <CreateMultiLineField saveButtonHandler={this.createNewField.bind(this)} groupName={this.props.group} fieldTypeOptions={options} cancelButtonHandler={this.props.closePanel} onFieldTypeChange={this.changeFieldType.bind(this)} />
                case FieldTypeKindEnum.Number:
                    return <CreateNumberField saveButtonHandler={this.createNewField.bind(this)} groupName={this.props.group} fieldTypeOptions={options} cancelButtonHandler={this.props.closePanel} onFieldTypeChange={this.changeFieldType.bind(this)} />
                case FieldTypeKindEnum.Currency:
                    return <CreateCurrencyField saveButtonHandler={this.createNewField.bind(this)} groupName={this.props.group} fieldTypeOptions={options} cancelButtonHandler={this.props.closePanel} onFieldTypeChange={this.changeFieldType.bind(this)} />
                case FieldTypeKindEnum.Choice:
                    return <CreateChoiceField saveButtonHandler={this.createNewField.bind(this)} groupName={this.props.group} fieldTypeOptions={options} cancelButtonHandler={this.props.closePanel} onFieldTypeChange={this.changeFieldType.bind(this)} />
                case FieldTypeKindEnum.Boolean:
                    return <CreateBooleanField saveButtonHandler={this.createNewField.bind(this)} groupName={this.props.group} fieldTypeOptions={options} cancelButtonHandler={this.props.closePanel} onFieldTypeChange={this.changeFieldType.bind(this)} />
                case FieldTypeKindEnum.URL:
                    return <CreateURLField saveButtonHandler={this.createNewField.bind(this)} groupName={this.props.group} fieldTypeOptions={options} cancelButtonHandler={this.props.closePanel} onFieldTypeChange={this.changeFieldType.bind(this)} />
                case FieldTypeKindEnum.DateTime:
                    return <CreateDateTimeField saveButtonHandler={this.createNewField.bind(this)} groupName={this.props.group} fieldTypeOptions={options} cancelButtonHandler={this.props.closePanel} onFieldTypeChange={this.changeFieldType.bind(this)} />
                default:
                    return null;
            }
        }

        return (
            <>
                { renderCreateFieldComponent(this.state.fieldType) }
            </>
        );
    }
}