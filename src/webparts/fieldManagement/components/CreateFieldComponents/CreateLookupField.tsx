import * as React from 'react';
import { PrimaryButton, Button, Dropdown, TextField, Toggle, IDropdownOption } from 'office-ui-fabric-react';
import { FieldTypeKindEnum } from '../FieldTypeKindEnum';
import { ISPField } from '../SPField';
import { ICreateFieldProps } from './ICreateFieldProps';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import { SPHttpClient } from '@microsoft/sp-http';

interface ICreateLookupFieldProps extends ICreateFieldProps {
    context: BaseComponentContext;
}

export const CreateLookupField: React.FC<ICreateLookupFieldProps> = (props) => {
    const [columnName, setColumnName] = React.useState("");
    const fieldType = FieldTypeKindEnum.Lookup;
    const [internalName, setInternalName] = React.useState("");
    const [group, setGroup] = React.useState(props.groupName);
    const [description, setDescription] = React.useState("");
    const [required, setRequired] = React.useState(false);
    const [enforceUniqueValues, setEnforceUniqueValues] = React.useState<boolean>(false);
    const [lists, setLists] = React.useState<IDropdownOption[]>([]);
    const [fields, setFields] = React.useState<IDropdownOption[]>([]);
    const [allowMultipleValues, setAllowMultipleValues] = React.useState<boolean>(false);
    const [allowUnlimitedLength, setAllowUnlimitedLength] = React.useState<boolean>(false);

    React.useEffect(() => {
        retrieveLists();
        return () => {
        }
    },
        []
    );

    const retrieveLists = () => {
        let context = props.context;
        let requestUrl = context.pageContext.web.absoluteUrl + `/_api/web/lists?&$filter=(TemplateFeatureId ne '00000000-0000-0000-0000-000000000000' and Hidden eq false)&$select=Id,Title`;
        context.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1).then((response) => {
            if (response.ok) {
                response.json().then((responseJSON) => {
                    console.log({responseJSON});
                    let lists: IDropdownOption[] = [];
                    responseJSON.value.forEach((list) => {
                        lists.push({ key: list.Id, text: list.Title });
                    });
                    setLists(lists);
                    retrieveFields(lists[0].key.toString());
                });
            }
        });
    }

    const retrieveFields = (guid: string) => {
        let context = props.context;
        let requestUrl = context.pageContext.web.absoluteUrl + `/_api/web/lists(guid'${guid}')/fields?&$filter=((FieldTypeKind eq 2 or FieldTypeKind eq 4 or FieldTypeKind eq 5 or FieldTypeKind eq 9) and Hidden eq false)&$select=InternalName,Title`;
        context.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1).then((response) => {
            if (response.ok) {
                response.json().then((responseJSON) => {
                    console.log({responseJSON});
                    let fields: IDropdownOption[] = [];
                    responseJSON.value.forEach((field) => {
                        fields.push({ key: field.InternalName, text: field.Title });
                    });
                    setFields(fields);
                });
            }
        });
    }

    const saveNewField = () => {
        let body: ISPField;

        body = {
            "@odata.type": "#SP.FieldLookup",
            Title: columnName,
            StaticName: internalName,
            InternalName: internalName,
            FieldTypeKind: fieldType,
            Required: required,
            EnforceUniqueValues: enforceUniqueValues,
            Group: group,
            Description: description,
            SchemaXml: `<Field Type="Lookup" Description="${description}" DisplayName="${columnName}" Required="${(required ? "TRUE" : "FALSE")}" EnforceUniqueValues="${(enforceUniqueValues ? "TRUE" : "FALSE")}" Group="${group}" StaticName="${internalName}" Name="${internalName}"></Field>`
        };

        props.saveButtonHandler(body);
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
            {
                lists.length == 0 ?
                    null :
                    <Dropdown label="Get Information From" defaultSelectedKey={lists[0].key} options={lists} />
            }
            {
                fields.length == 0 ?
                    null :
                    <Dropdown label="In This Column" defaultSelectedKey={fields[0].key} options={fields} />
            }
            <Toggle label="Allow Multiple Values" onChanged={(evt) => setAllowMultipleValues(evt)} />
            <Toggle label="Allow Unlimited Length in Document Libraries" onChanged={(evt) => setAllowUnlimitedLength(evt)} />
            <br />
            <PrimaryButton text="Save" onClick={() => saveNewField()} />
            <Button text="Cancel" onClick={() => props.cancelButtonHandler()} />
        </>
    );
}