import * as React from 'react';
import { PrimaryButton, Button, Dropdown, IDropdownOption, TextField, Toggle } from 'office-ui-fabric-react';
import { FieldTypeKindEnum } from '../FieldTypeKindEnum';
import { ISPField } from '../SPField';

export interface CreateCurrencyFieldProps {
    fieldTypeOptions: IDropdownOption[];
    saveButtonHandler: Function;
    cancelButtonHandler: Function;
    groupName: string;
    onFieldTypeChange: Function;
}

export const CreateCurrencyField: React.FC<CreateCurrencyFieldProps> = (props) => {
    const [columnName, setColumnName] = React.useState("");
    const [fieldType, setFieldType] = React.useState(FieldTypeKindEnum.Currency);
    const [internalName, setInternalName] = React.useState("");
    const [group, setGroup] = React.useState(props.groupName);
    const [description, setDescription] = React.useState("");
    const [required,setRequired] = React.useState(false);
    const [enforceUniqueValues, setEnforceUniqueValues] = React.useState<boolean>(false);
    const [defaultValue, setDefaultValue] = React.useState();
    const [minValue, setMinValue] = React.useState();
    const [maxValue, setMaxValue] = React.useState();
    const [displayFormat, setDisplayFormat] = React.useState(-1);
    const [selectedCurrency, setSelectedCurrency] = React.useState("1033");

    const saveNewField = () => {
        let body: ISPField;
        let defaultString = defaultValue.length == 0 ? '' : `<Default>${defaultValue}</Default>`;
        let minString = minValue == null ? '' : 'Min="' + minValue + '"';
        let maxString = maxValue == null ? '' : 'Max="' + maxValue + '"';
        
        body = {
            "@odata.type": "#SP.FieldCurrency",
            Title: columnName,
            StaticName: internalName,
            InternalName: internalName,
            FieldTypeKind: FieldTypeKindEnum.Currency,
            Required: required,
            EnforceUniqueValues: enforceUniqueValues,
            DefaultValue: defaultValue,
            Group: group,
            DisplayFormat: +(displayFormat),
            Description: description,
            SchemaXml: `<Field Type="Currency" 
            Description="${description}" 
            DisplayName="${columnName}" 
            Required="${(required? "TRUE" : "FALSE")}" 
            EnforceUniqueValues="${(enforceUniqueValues? "TRUE" : "FALSE")}" 
            Group="${group}" 
            StaticName="${internalName}" 
            Decimals="${displayFormat}" 
            LCID="${selectedCurrency}" 
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
                <Dropdown label="Currency" defaultSelectedKey={selectedCurrency} options={currencyOptions} onChanged={(evt: any) => {
                    setSelectedCurrency(evt.key);
                }} />
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
                <br />
                <PrimaryButton text="Save" onClick={() => saveNewField()} />
                <Button text="Cancel" onClick={() => props.cancelButtonHandler()} />
            </>
    );
}