import * as React from 'react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { PrimaryButton, Button, Dropdown, IDropdownOption, FacepileBase, IChoiceGroupOption, ChoiceGroup, IDropdown, DatePicker, DayOfWeek } from 'office-ui-fabric-react';
import { ISPHttpClientOptions, SPHttpClientResponse, SPHttpClient } from '@microsoft/sp-http';
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
    displayDefaultValueDTInput: boolean
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
        let minString: string = '';
        let maxString: string = '';
        let defaultString: string = '';
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
                minString = data.minValue == null ? '' : (data.showAsPercentage ? 'Min="' + data.minValue/100 + '"' : 'Min="' + data.minValue + '"');
                maxString = data.maxValue == null ? '' : (data.showAsPercentage ? 'Max="' + data.maxValue/100 + '"' : 'Max="' + data.maxValue + '"');
                defaultString = data.defaultValue.length == 0 ? '' : "<Default>" + (data.showAsPercentage ? (+(data.defaultValue)/100).toString() : data.defaultValue) + "</Default>";
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
            case FieldTypeKindEnum.Currency:
                    minString = data.minValue == null ? '' : (data.showAsPercentage ? 'Min="' + data.minValue/100 + '"' : 'Min="' + data.minValue + '"');
                    maxString = data.maxValue == null ? '' : (data.showAsPercentage ? 'Max="' + data.maxValue/100 + '"' : 'Max="' + data.maxValue + '"');
                    defaultString = data.defaultValue.length == 0 ? '' : "<Default>" + (data.showAsPercentage ? (+(data.defaultValue)/100).toString() : data.defaultValue) + "</Default>";
                    body = {
                        "@odata.type": "#SP.FieldCurrency",
                        Title: data.columnName,
                        StaticName: data.internalName,
                        InternalName: data.internalName,
                        FieldTypeKind: FieldTypeKindEnum.Currency,
                        Required: data.required,
                        EnforceUniqueValues: data.enforceUniqueValues,
                        DefaultValue: data.defaultValue,
                        Group: data.group,
                        DisplayFormat: +(data.displayFormat),
                        Description: data.description,
                        SchemaXml: '<Field Type="Currency" DisplayName="'+ data.columnName + '" LCID="'+data.selectedCurrency+'" Description="'+data.description+'" Required="'+ (data.required? "TRUE" : "FALSE") +'" EnforceUniqueValues="'+ (data.enforceUniqueValues? "TRUE" : "FALSE") +'" Decimals="'+data.displayFormat+'" Group="'+data.group+'" StaticName="'+data.internalName+'" Name="'+data.internalName+'" Version="1" '+ minString + ' ' + maxString + '>'+ defaultString +'</Field>'
                    }
            break;
            case FieldTypeKindEnum.Boolean:
                    body = {
                        "@odata.type": "#SP.Field",
                        Title: data.columnName,
                        StaticName: data.internalName,
                        InternalName: data.internalName,
                        FieldTypeKind: FieldTypeKindEnum.Boolean,
                        Group: data.group,
                        Description: data.description,
                        SchemaXml: '<Field Type="Boolean" DisplayName="'+ data.columnName + '" Required="FALSE" EnforceUniqueValues="FALSE" Description="'+data.description+'" Group="'+data.group+'" StaticName="'+data.internalName+'" Name="'+data.internalName+'"><Default>'+data.defaultBooleanValueAsString+'</Default></Field>'
                    }
            break;
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
                    }
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
            { key: FieldTypeKindEnum.Currency , text: 'Currency ($, ¥, €)' },
            { key: FieldTypeKindEnum.DateTime , text: 'Date and Time' },
            { key: FieldTypeKindEnum.Lookup , text: 'Lookup (information already on this site)', disabled: true },
            { key: FieldTypeKindEnum.Boolean , text: 'Yes/No (check box)' },
            { key: FieldTypeKindEnum.User , text: 'Person or Group', disabled: true },
            { key: FieldTypeKindEnum.URL , text: 'Hyperlink or Picture' },
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

          const currencyOptions: IDropdownOption[] = [
            {key: "1164", text: "؋123,456.00 (Afghanistan)"}, 
            {key: "1052", text: "123,456.00 Lekë (Albania)"}, 
            {key: "5121", text: "123,456.00 د.ج. (Algeria)"}, 
            {key: "11274", text: "$ 123,456.00 (Argentina)"}, 
            {key: "1067", text: "123,456.00 ֏ (Armenia)"}, 
            {key: "3081", text: "$123,456.00 (Australia)"}, 
            {key: "3079", text: "€ 123,456.00 (Austria)"}, 
            {key: "1068", text: "₼ 123,456.00 (Azerbaijan)"}, 
            {key: "2092", text: "123,456.00 ₼ (Azerbaijan)"}, 
            {key: "15361", text: "123,456.00 د.ب. (Bahrain)"}, 
            {key: "2117", text: "123,456.00৳ (Bangladesh)"}, 
            {key: "1133", text: "123,456.00 ₽ (Bashkir)"}, 
            {key: "1059", text: "123,456.00 Br (Belarus)"}, 
            {key: "2067", text: "123,456.00 € (Belgium)"}, 
            {key: "10249", text: "$123,456.00 (Belize)"}, 
            {key: "16394", text: "Bs123,456.00 (Bolivia)"}, 
            {key: "8218", text: "123,456.00 КМ (Bosnia and Herzegovina)"}, 
            {key: "5146", text: "123,456.00 KM (Bosnia and Herzegovina)"}, 
            {key: "1046", text: "R$123,456.00 (Brazil)"}, 
            {key: "2110", text: "$ 123,456.00 (Brunei Darussalam)"}, 
            {key: "1026", text: "123,456.00 лв. (Bulgaria)"}, 
            {key: "1107", text: "123,456.00៛ (Cambodia)"}, 
            {key: "3084", text: "123,456.00 $ (Canada)"}, 
            {key: "4105", text: "$123,456.00 (Canada)"}, 
            {key: "13322", text: "$123,456.00 (Chile)"}, 
            {key: "9226", text: "$ 123,456.00 (Columbia)"}, 
            {key: "5130", text: "₡123,456.00 (Costa Rica)"}, 
            {key: "1050", text: "123,456.00 kn (Croatia)"}, 
            {key: "1029", text: "123,456.00 Kč (Czech Republic)"}, 
            {key: "1030", text: "123,456.00 kr. (Denmark)"}, 
            {key: "7178", text: "$123,456.00 (Dominican Republic)"}, 
            {key: "12298", text: "$123,456.00 (Ecuador)"}, 
            {key: "3073", text: "123,456.00 ج.م. (Egypt)"}, 
            {key: "17418", text: "$123,456.00 (El Salvador)"}, 
            {key: "1061", text: "123,456.00 € (Estonia)"}, 
            {key: "1118", text: "ብር123,456.00 (Ethiopia)"}, 
            {key: "1080", text: "123,456.00 kr (Faroe Islands)"}, 
            {key: "1035", text: "123,456.00 € (Finland)"}, 
            {key: "1036", text: "123,456.00 € (France)"}, 
            {key: "1079", text: "123,456.00 ₾ (Georgia)"}, 
            {key: "1031", text: "123,456.00 € (Germany)"}, 
            {key: "1032", text: "123,456.00 € (Greece)"}, 
            {key: "1135", text: "kr.123,456.00 (Greenland)"}, 
            {key: "4106", text: "Q123,456.00 (Guatemala)"}, 
            {key: "18442", text: "L123,456.00 (Honduras)"}, 
            {key: "3076", text: "HK$123,456.00 (Hong Kong S.A.R.)"}, 
            {key: "1038", text: "123,456.00 Ft (Hungary)"}, 
            {key: "1039", text: "123,456.00 ISK (Iceland)"}, 
            {key: "1081", text: "₹123,456.00 (India)"}, 
            {key: "1057", text: "Rp123,456.00 (Indonesia)"}, 
            {key: "1065", text: "123,456.00ريال (Iran)"}, 
            {key: "2049", text: "123,456.00 د.ع. (Iraq)"}, 
            {key: "6153", text: "€123,456.00 (Ireland)"}, 
            {key: "1040", text: "123,456.00 € (Italy)"}, 
            {key: "1037", text: "₪ 123,456.00 (Israel)"}, 
            {key: "8201", text: "$123,456.00 (Jamaica)"}, 
            {key: "1041", text: "¥123,456.00 (Japan)"}, 
            {key: "11265", text: "123,456.00 د.ا. (Jordan)"}, 
            {key: "1087", text: "123,456.00 ₸ (Kazakhstan)"}, 
            {key: "1089", text: "Ksh123,456.00 (Kenya)"}, 
            {key: "1042", text: "₩123,456.00 (Korea)"}, 
            {key: "13313", text: "123,456.00 د.ك. (Kuwait)"}, 
            {key: "1088", text: "123,456.00 сом (Kyrgyzstan)"}, 
            {key: "1108", text: "₭123,456.00 (Lao P.D.R)"}, 
            {key: "1062", text: "123,456.00 € (Latvia)"}, 
            {key: "12289", text: "123,456.00 ل.ل. (Lebanon)"}, 
            {key: "4097", text: "123,456.00 د.ل. (Libya)"}, 
            {key: "5127", text: "CHF 123,456.00 (Liechtenstein)"}, 
            {key: "1063", text: "123,456.00 € (Lithuania)"}, 
            {key: "5132", text: "123,456.00 € (Luxembourg)"}, 
            {key: "5124", text: "MOP123,456.00 (Macao S.A.R.)"}, 
            {key: "1071", text: "ден 123,456.00 (North Macedonia)"}, 
            {key: "1086", text: "RM123,456.00 (Malaysia)"}, 
            {key: "1125", text: "123,456.00 ރ. (Maldives)"}, 
            {key: "1082", text: "€123,456.00 (Malta)"}, 
            {key: "2058", text: "$123,456.00 (Mexico)"}, 
            {key: "6156", text: "123,456.00 € (Monaco)"}, 
            {key: "1104", text: "₮ 123,456.00 (Mongolia)"}, 
            {key: "6145", text: "123,456.00 د.م. (Morocco)"}, 
            {key: "1121", text: "रु 123,456.00 (Nepal)"}, 
            {key: "1043", text: "€ 123,456.00 (Netherlands)"}, 
            {key: "5129", text: "$123,456.00 (New Zealand)"}, 
            {key: "19466", text: "C$123,456.00 (Nicaragua)"}, 
            {key: "1128", text: "₦ 123,456.00 (Nigeria)"}, 
            {key: "1044", text: "kr 123,456.00 (Norway)"}, 
            {key: "8193", text: "123,456.00 ر.ع. (Oman)"}, 
            {key: "1056", text: "Rs 123,456.00 (Pakistan)"}, 
            {key: "6154", text: "B/.123,456.00 (Panama)"}, 
            {key: "15370", text: "₲ 123,456.00 (Paraguay)"}, 
            {key: "2052", text: "¥123,456.00 (People's Republic of China)"}, 
            {key: "10250", text: "S/123,456.00 (Peru)"}, 
            {key: "13321", text: "₱123,456.00 (Philippines)"}, 
            {key: "1045", text: "123,456.00 zł (Poland)"}, 
            {key: "2070", text: "123,456.00 € (Portugal)"}, 
            {key: "20490", text: "$123,456.00 (Puerto Rico)"}, 
            {key: "16385", text: "123,456.00 ر.ق. (Qatar)"}, 
            {key: "1048", text: "123,456.00 RON (Romania)"}, 
            {key: "1049", text: "123,456.00 ₽ (Russia)"}, 
            {key: "1159", text: "RF 123,456.00 (Rwanda)"}, 
            {key: "1025", text: "123,456.00 ر.س. (Saudi Arabia)"}, 
            {key: "1160", text: "123,456.00 CFA (Senegal)"}, 
            {key: "9242", text: "123,456.00 RSD (Serbia, Latin)"}, 
            {key: "12314", text: "123,456.00 € (Montenegro)"}, 
            {key: "10266", text: "123,456.00 дин. (Serbia, Cyrillic)"}, 
            {key: "4100", text: "$123,456.00 (Singapore)"}, 
            {key: "1051", text: "123,456.00 € (Slovakia)"}, 
            {key: "1060", text: "123,456.00 € (Slovenia)"}, 
            {key: "7177", text: "R123,456.00 (South Africa)"}, 
            {key: "3082", text: "123,456.00 € (Spain)"}, 
            {key: "1115", text: "රු.123,456.00 (Sri Lanka)"}, 
            {key: "1053", text: "123,456.00 kr (Sweden)"}, 
            {key: "2055", text: "CHF 123,456.00 (Switzerland)"}, 
            {key: "10241", text: "123,456.00 ل.س. (Syria)"}, 
            {key: "1028", text: "NT$123,456.00 (Taiwan)"}, 
            {key: "1064", text: "123,456.00 смн (Tajikistan)"}, 
            {key: "1054", text: "฿123,456.00 (Thailand)"}, 
            {key: "11273", text: "$123,456.00 (Trinidad and Tobago)"}, 
            {key: "7169", text: "123,456.00 د.ت. (Tunisia)"}, 
            {key: "1055", text: "123,456.00 ₺ (Turkey)"}, 
            {key: "1090", text: "123,456.00m. (Turkmenistan)"}, 
            {key: "1058", text: "123,456.00 ₴ (Ukraine)"}, 
            {key: "14337", text: "123,456.00 د.إ. (United Arab Emirates)"}, 
            {key: "2057", text: "£123,456.00 (United Kingdom)"}, 
            {key: "1033", text: "$123,456.00 (United States)"}, 
            {key: "14346", text: "$ 123,456.00 (Uruguay)"}, 
            {key: "1091", text: "123,456.00 soʻm (Uzbekistan)"}, 
            {key: "2115", text: "сўм 123,456.00 (Uzbekistan)"}, 
            {key: "8202", text: "Bs.S123,456.00 (Venezuela)"}, 
            {key: "1066", text: "123,456.00 ₫ (Vietnam)"}, 
            {key: "9217", text: "123,456.00 ر.ي. (Yemen)"}, 
            {key: "12297", text: "US$123,456.00 (Zimbabwe)"}, 
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
                { 
                    this.state.fieldType == FieldTypeKindEnum.Text ?
                        <>
                            <Toggle label="Required" onChanged={(evt) => this.setState({required: evt})} />
                            <Toggle label="Enforce Unique Values" onChanged={(evt) => this.setState({enforceUniqueValues: evt})} />
                            <TextField label="Maximum number of characters" max={255} min={0} type="number" defaultValue="255" onChange={(evt: React.FormEvent<HTMLInputElement>) => { this.setState({ maxNoCharacters: +((evt.target as any).value) })}} />
                            <TextField label="Default value" value={this.state.defaultValue} onChange={(evt: React.FormEvent<HTMLInputElement>) => { this.setState({ defaultValue: (evt.target as any).value })}} />
                        </> : null
                }
                {
                    this.state.fieldType == FieldTypeKindEnum.Note ?
                        <>
                            <Toggle label="Required" onChanged={(evt) => this.setState({required: evt})} />
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
                            <Toggle label="Required" onChanged={(evt) => this.setState({required: evt})} />
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
                            <Toggle label="Required" onChanged={(evt) => this.setState({required: evt})} />
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
                {
                    this.state.fieldType == FieldTypeKindEnum.Currency ?
                        <>
                            <Toggle label="Required" onChanged={(evt) => this.setState({required: evt})} />
                            <Toggle label="Enforce Unique Values" onChanged={(evt) => this.setState({enforceUniqueValues: evt})} />
                            <Dropdown label="Currency" defaultSelectedKey={this.state.selectedCurrency} options={currencyOptions} onChanged={(evt: any) => {
                                this.setState({selectedCurrency: evt.key})
                            }} />
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
                        </> : null
                }
                {
                    this.state.fieldType == FieldTypeKindEnum.Boolean ?
                    <>
                        <Dropdown label="Default value" options={[{key: '1', text: 'Yes'}, {key: '0', text: 'No'}]} defaultSelectedKey={this.state.defaultBooleanValueAsString} onChanged={(evt:any) => {
                            this.setState({defaultBooleanValueAsString: evt.key})
                        }} />
                    </>
                    : null
                }
                {
                    this.state.fieldType == FieldTypeKindEnum.URL ? 
                    <>
                        <Toggle label="Required" onChanged={(evt) => this.setState({required: evt})} />
                        <Dropdown label="Format URL as" options={[{key: 'Hyperlink', text: 'Hyperlink'}, {key: 'Image', text: 'Picture'}]} defaultSelectedKey={this.state.urlFieldFormat} onChanged={(evt:any) => {
                            this.setState({urlFieldFormat: evt.key})
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
                                this.setState({dateAndTimeFormat: evt.key})
                            }} />
                        <ChoiceGroup styles={{flexContainer: {display: "flex"}}} label="Display Format" defaultSelectedKey={this.state.friendlyDisplayFormat} options={[{key: "Disabled", text: "Standard\u00A0\u00A0"},{key: "Relative", text: "Friendly"}]} onChanged={(evt: any) => { 
                                this.setState({friendlyDisplayFormat: evt.key})
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
                                this.setState({dateAndTimeFormat: evt.key})
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
    };

    private _onSelectDate = (date: Date | null | undefined): void => {
        //Need to compensate the time difference between GMT and LocaleTime
        this.setState({defaultValue: (new Date(date.getTime() + Math.abs(date.getTimezoneOffset() * (-60000))).toISOString())});
    };
}