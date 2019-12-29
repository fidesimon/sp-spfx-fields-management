import * as React from 'react';
import { PrimaryButton, Button, Dropdown, TextField, Toggle, ChoiceGroup, DatePicker, DayOfWeek, IChoiceGroupOption, IDropdownOption, concatStyleSets } from 'office-ui-fabric-react';
import { FieldTypeKindEnum } from '../FieldTypeKindEnum';
import { ISPField } from '../SPField';
import { ICreateFieldProps } from './ICreateFieldProps';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import { SPHttpClientResponse, SPHttpClient } from '@microsoft/sp-http';

interface ICreateUserFieldProps extends ICreateFieldProps{
    context: BaseComponentContext;
}

export const CreateUserField: React.FC<ICreateUserFieldProps> = (props) => {
    const [columnName, setColumnName] = React.useState("");
    const fieldType = FieldTypeKindEnum.User;
    const [internalName, setInternalName] = React.useState("");
    const [group, setGroup] = React.useState(props.groupName);
    const [description, setDescription] = React.useState("");
    const [required, setRequired] = React.useState(false);
    const [enforceUniqueValues, setEnforceUniqueValues] = React.useState<boolean>(false);
    const [allowMultipleSelection, setAllowMultipleSelection] = React.useState<boolean>(false);
    const [userSelectionMode, setUserSelectionMode] = React.useState<string>("PeopleOnly");
    const [showField, setShowField] = React.useState("ImnName");
    const [chooseFrom, setChooseFrom] = React.useState("0");
    const [groupFields, setGroupFields] = React.useState<IDropdownOption[]>([]);
    const [selectedGroup, setSelectedGroup] = React.useState<string>("");

    React.useEffect(() => {
        _retrieveColumns();
        return () => {
        }
      },
      []
    );

    const saveNewField = () => {
        let body: ISPField;
        let multi: string = allowMultipleSelection ? 'Mult="TRUE"' : '';

        body = {
            "@odata.type": "#SP.FieldUser",
            Title: columnName,
            StaticName: internalName,
            InternalName: internalName,
            FieldTypeKind: fieldType,
            Required: required,
            EnforceUniqueValues: enforceUniqueValues,
            Group: group,
            Description: description,
            SchemaXml: `<Field Type="${allowMultipleSelection ? "UserMulti" : "User"}" ${multi} ShowField="${showField}" Description="${description}" UserSelectionMode="${userSelectionMode}" UserSelectionScope="${chooseFrom === "0" ? "0" : selectedGroup}" DisplayName="${columnName}" List="UserInfo" Required="${(required ? "TRUE" : "FALSE")}" EnforceUniqueValues="${(enforceUniqueValues ? "TRUE" : "FALSE")}" Group="${group}" StaticName="${internalName}" Name="${internalName}"></Field>`
        };

        props.saveButtonHandler(body);
    }

    const allowSelectionOfChoices: IChoiceGroupOption[] = [
        { key: "PeopleOnly", text: "People Only\u00A0\u00A0" },
        { key: "PeopleAndGroups", text: "People and Groups" }
    ];

    const chooseFromChoices: IChoiceGroupOption[] = [
        { key: "0", text: "All Users" },
        { key: "1", text: "SharePoint Group" }
    ];

    const showFieldDropDown: IDropdownOption[] = [
        { key: "Title", text: "Name" },
        { key: "ComplianceAssetId", text: "Compliance Asset Id" },
        { key: "Name", text: "Account" },
        { key: "EMail", text: "Work email" },
        { key: "OtherMail", text: "OtherMail" },
        { key: "UserExpiration", text: "User Expiration" },
        { key: "UserLastDeletionTime", text: "User Last Deletion Time" },
        { key: "MobilePhone", text: "Mobile phone" },
        { key: "SipAddress", text: "SIP Address" },
        { key: "Department", text: "Department" },
        { key: "JobTitle", text: "Title" },
        { key: "FirstName", text: "First name" },
        { key: "LastName", text: "Last name" },
        { key: "WorkPhone", text: "Work phone" },
        { key: "UserName", text: "User name" },
        { key: "Office", text: "Office" },
        { key: "ID", text: "ID" },
        { key: "Modified", text: "Modified" },
        { key: "Created", text: "Created" },
        { key: "ImnName", text: "Name (with presence)" },
        { key: "PictureOnly_Size_36px", text: "Picture Only (36x36)" },
        { key: "PictureOnly_Size_48px", text: "Picture Only (48x48)" },
        { key: "PictureOnly_Size_72px", text: "Picture Only (72x72)" },
        { key: "NameWithPictureAndDetails", text: "Name (with picture and details)" },
        { key: "ContentTypeDisp", text: "Content Type" }
    ];

    const _retrieveColumns = () => {
        let context = props.context;
        let requestUrl = context.pageContext.web.absoluteUrl + `/_api/web/sitegroups?&$select=Title,Id`;
        context.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1).then((response) => {
            if(response.ok){
            response.json().then((responseJSON) => {
                let siteGroups: IDropdownOption[] = [];
                responseJSON.value.forEach((item) => {
                    siteGroups.push({ key: item.Id.toString(), text: item.Title });
                });
                setGroupFields(siteGroups);
                setSelectedGroup(siteGroups[0].key.toString());
            });
            }
        });
      }

    return (
        <>
            <TextField label="Column Name" required value={columnName} onChanged={(evt) => {
                setColumnName(evt);
                let newValue = evt.replace(/[^A-Z0-9]+/ig, "");
                setInternalName(newValue.length >= 32 ? newValue.substr(0, 32) : newValue);
            }} />
            <Dropdown label="Field Type" options={props.fieldTypeOptions} defaultSelectedKey={fieldType} onChanged={(evt: any) => { props.onFieldTypeChange(evt.key) }} />
            <TextField label="Internal Name" required value={internalName} onChanged={(evt) => { setInternalName(evt) }} />
            <TextField label="Group" defaultValue={group} onChange={(evt: React.FormEvent<HTMLInputElement>) => { setGroup((evt.target as any).value); }} />
            <TextField label="Description" multiline autoAdjustHeight onChange={(evt: React.FormEvent<HTMLTextAreaElement>) => { setDescription((evt.target as any).value) }} />
            <Toggle label="Required" onChanged={(evt) => setRequired(evt)} />
            <Toggle label="Enforce Unique Values" onChanged={(evt) => setEnforceUniqueValues(evt)} />
            <Toggle label="Allow Multiple Selection" onChanged={(evt) => setAllowMultipleSelection(evt)} />
            <ChoiceGroup styles={{ flexContainer: { display: "flex" } }} label="Allow Selection of" defaultSelectedKey={userSelectionMode} options={allowSelectionOfChoices} onChanged={(evt: any) => {
                setUserSelectionMode(evt.key);
            }} />
            <ChoiceGroup label="Choose From" defaultSelectedKey={chooseFrom} options={chooseFromChoices} onChanged={(evt: any) => {
                setChooseFrom(evt.key);
            }} />
            {groupFields.length === 0 ? 
                <Dropdown disabled={chooseFrom === "0"} options={groupFields} onChanged={(evt:any) => setSelectedGroup(evt.key)} />:
                <Dropdown disabled={chooseFrom === "0"} options={groupFields} defaultSelectedKey={groupFields[0].key} onChanged={(evt:any) => setSelectedGroup(evt.key)} />
            }
            <Dropdown label="Show Field" defaultSelectedKey={showField} options={showFieldDropDown} onChanged={(evt: any) => {
                setShowField(evt.key);
            }} />
            <br />
            <PrimaryButton text="Save" onClick={() => saveNewField()} />
            <Button text="Cancel" onClick={() => props.cancelButtonHandler()} />
        </>
    );
}