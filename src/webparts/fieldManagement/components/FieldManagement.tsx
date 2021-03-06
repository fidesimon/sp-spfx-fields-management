import * as React from 'react';
import styles from './FieldManagement.module.scss';
import { IFieldManagementProps } from './IFieldManagementProps';
import { Panel, PanelType, DetailsList, GroupedList, Pivot, PivotItem, PivotLinkFormat, ICommandBarItemProps, CommandBar, Layer, Separator, createTheme, ITheme, getTheme, Label, Spinner } from 'office-ui-fabric-react';
import { IGroup, Group } from './Group';
import { GroupList } from './GroupList';
import FieldDisplay from './FieldDisplay';
import FieldCreate from './FieldCreate';
import DisplayGroups from './FieldsList/DisplayGroups';
import GroupHeader from './FieldsList/GroupHeader';

import 'office-ui-fabric-core/dist/css/fabric.css';

import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions
} from '@microsoft/sp-http';
import { ISPField } from './SPField';

export interface IFieldManagementState {
  ListOfGroups: IGroup[];
  isPanelOpened: boolean;
  fieldToDisplay: ISPField;
  fieldsPlain: DataDetailedList;
  isCreateFieldPanelOpen: boolean;
  createFieldGroupName: string;
  fieldCreationFunction: Function;
}

export class groupsDetailedList {
  key: string;
  name: string;
  startIndex: number;
  count: number;
}

export class DataDetailedList {
  groups: groupsDetailedList[];
  items: ISPField[];
}

const theme: ITheme = createTheme({
  fonts: {
    medium: {
      fontSize: '18px'
    }
  }
});

export default class FieldManagement extends React.Component<IFieldManagementProps, IFieldManagementState> {
  constructor(props) {
    super(props);
    //this.state = { ListOfGroups: [], createFieldGroupName: '', isPanelOpened: false, isCreateFieldPanelOpen: false, fieldToDisplay: null, fieldsPlain: this.magicWithGroups(this.mockData) };
    this.state = { ListOfGroups: this.mockData, createFieldGroupName: '', isPanelOpened: false, isCreateFieldPanelOpen: false, fieldToDisplay: null, fieldsPlain: this.magicWithGroups(this.mockData), fieldCreationFunction: null };
    this.deleteField.bind(this);
  }

  magicWithGroups(groups: IGroup[]): DataDetailedList {
    let groupz: groupsDetailedList[] = [];
    let items: ISPField[] = [];
    let count: number[] = [];
    let itemCounter: number = 0;
    for (let i = 0; i < groups.length; i++) {
      items.push(...groups[i].Fields);
      count.push(groups[i].Fields.length);
      groupz.push({ key: groups[i].Name, name: groups[i].Name, startIndex: itemCounter, count: count[i] });
      itemCounter += count[i];
    }
    return {
      groups: groupz,
      items: items
    };
  }

  mockData: Array<IGroup> = [
    {
      Name: "AAA", Ascending: true, Fields: [{
        AutoIndexed: false,
        CanBeDeleted: true,
        ClientSideComponentId: "00000000-0000-0000-0000-000000000000",
        ClientSideComponentProperties: null,
        ClientValidationFormula: null,
        ClientValidationMessage: null,
        CustomFormatter: null,
        DefaultFormula: null,
        DefaultValue: null,
        Description: "",
        Direction: "none",
        EnforceUniqueValues: false,
        EntityPropertyName: "OrganizationalIDNumber",
        FieldTypeKind: 2,
        Filterable: true,
        FromBaseType: false,
        Group: "Core Contact and Calendar Columns",
        Hidden: false,
        Id: "0850ae15-19dd-431f-9c2f-3aff3ae292ce",
        IndexStatus: 0,
        Indexed: false,
        InternalName: "OrganizationalIDNumber",
        JSLink: "clienttemplates.js",
        MaxLength: 255,
        PinnedToFiltersPane: false,
        ReadOnlyField: false,
        Required: false,
        SchemaXml: '<Field ID="{0850AE15-19DD-431f-9C2F-3AFF3AE292CE}" Name="OrganizationalIDNumber" StaticName="OrganizationalIDNumber" SourceID="http://schemas.microsoft.com/sharepoint/v3" DisplayName="Organizational ID Number" Group="Core Contact and Calendar Columns" Type="Text" Sealed="TRUE" AllowDeletion="TRUE" />',
        Scope: "/sites/firstTest",
        Sealed: true,
        ShowInFiltersPane: 0,
        Sortable: true,
        StaticName: "OrganizationalIDNumber",
        Title: "Organizational ID Number",
        TypeAsString: "Text",
        TypeDisplayName: "Single line of text",
        TypeShortDescription: "Single line of text",
        ValidationFormula: null,
        ValidationMessage: null
      },
      {
        AutoIndexed: false,
        CanBeDeleted: true,
        ClientSideComponentId: "00000000-0000-0000-0000-000000000000",
        ClientSideComponentProperties: null,
        ClientValidationFormula: null,
        ClientValidationMessage: null,
        CustomFormatter: null,
        DefaultFormula: null,
        DefaultValue: null,
        Description: "References to resources from which this resource was derived",
        Direction: "none",
        EnforceUniqueValues: false,
        EntityPropertyName: "OData__Source",
        FieldTypeKind: 3,
        Filterable: false,
        FromBaseType: false,
        Group: "Core Document Columns",
        Hidden: false,
        Id: "b0a3c1db-faf1-48f0-9be1-47d2fc8cb5d6",
        IndexStatus: 0,
        Indexed: false,
        InternalName: "_Source",
        JSLink: "clienttemplates.js",
        PinnedToFiltersPane: false,
        ReadOnlyField: false,
        Required: false,
        SchemaXml: '<Field ID="{B0A3C1DB-FAF1-48f0-9BE1-47D2FC8CB5D6}" Type="Note" NumLines="2" Group="Core Document Columns" Name="_Source" DisplayName="Source" SourceID="http://schemas.microsoft.com/sharepoint/v3/fields" StaticName="_Source" Description="References to resources from which this resource was derived" />',
        Scope: "/sites/firstTest",
        Sealed: false,
        ShowInFiltersPane: 0,
        Sortable: false,
        StaticName: "_Source",
        Title: "Source",
        TypeAsString: "Note",
        TypeDisplayName: "Multiple lines of text",
        TypeShortDescription: "Multiple lines of text",
        ValidationFormula: null,
        ValidationMessage: null
      }]
    },
    {
      Name: "BBB", Ascending: true, Fields: [{
        AutoIndexed: false,
        CanBeDeleted: false,
        ClientSideComponentId: "00000000-0000-0000-0000-000000000000",
        ClientSideComponentProperties: null,
        ClientValidationFormula: null,
        ClientValidationMessage: null,
        CustomFormatter: null,
        DefaultFormula: null,
        DefaultValue: null,
        Description: "",
        Direction: "none",
        EnforceUniqueValues: false,
        EntityPropertyName: "Group",
        FieldTypeKind: 1,
        Filterable: true,
        FromBaseType: false,
        Group: "Base Columns",
        Hidden: false,
        Id: "c86a2f7f-7680-4a0b-8907-39c4f4855a35",
        IndexStatus: 0,
        Indexed: false,
        InternalName: "Group",
        JSLink: "clienttemplates.js",
        PinnedToFiltersPane: false,
        ReadOnlyField: false,
        Required: false,
        SchemaXml: '<Field ID="{c86a2f7f-7680-4a0b-8907-39c4f4855a35}" Name="Group" SourceID="http://schemas.microsoft.com/sharepoint/v3" StaticName="Group" Group="Base Columns" DisplayName="Group Type" Required="FALSE" Sealed="TRUE" Type="Integer" />',
        Scope: "/sites/firstTest",
        Sealed: true,
        ShowInFiltersPane: 0,
        Sortable: true,
        StaticName: "Group",
        Title: "Group Type",
        TypeAsString: "Integer",
        TypeDisplayName: "Integer",
        TypeShortDescription: "Integer",
        ValidationFormula: null,
        ValidationMessage: null
      }]
    },
    {
      Name: "CCC", Ascending: true, Fields: [{
        AutoIndexed: false,
        CanBeDeleted: true,
        Choices: ["Lorem", "Ipsum", "Sit", "Mit", "Dolor"],
        ClientSideComponentId: "00000000-0000-0000-0000-000000000000",
        ClientSideComponentProperties: null,
        ClientValidationFormula: null,
        ClientValidationMessage: null,
        CustomFormatter: null,
        DefaultFormula: null,
        DefaultValue: "Lorem",
        Description: "",
        Direction: "none",
        EnforceUniqueValues: false,
        EntityPropertyName: "TestField",
        FieldTypeKind: 6,
        Filterable: true,
        FromBaseType: false,
        Group: "Szymon Fields",
        Hidden: false,
        Id: "869cad7e-d92b-4938-838e-c1201c31b7c4",
        IndexStatus: 0,
        Indexed: false,
        InternalName: "TestField",
        JSLink: "clienttemplates.js",
        PinnedToFiltersPane: false,
        ReadOnlyField: false,
        Required: false,
        SchemaXml: '<Field Type="Choice" DisplayName="TestField" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" FillInChoice="FALSE" Group="Szymon Fields" ID="{869cad7e-d92b-4938-838e-c1201c31b7c4}" SourceID="{210bdfec-81e1-4b01-b5a5-571ae3ce1cc6}" StaticName="TestField" Name="TestField" Version="1"><Default>Lorem</Default><CHOICES><CHOICE>Lorem</CHOICE><CHOICE>Ipsum</CHOICE><CHOICE>Sit</CHOICE><CHOICE>Mit</CHOICE><CHOICE>Dolor</CHOICE></CHOICES></Field>',
        Scope: "/sites/firstTest",
        Sealed: false,
        ShowInFiltersPane: 0,
        Sortable: true,
        StaticName: "TestField",
        Title: "TestField",
        TypeAsString: "Choice",
        TypeDisplayName: "Choice",
        TypeShortDescription: "Choice (menu to choose from)",
        ValidationFormula: null,
        ValidationMessage: null
      }]
    },

  ];

  componentWillMount() {
    const grupki = this._retrieveColumns();
  }

  handleFieldClick = (fieldData: ISPField) => {
    this.setState({ isPanelOpened: true, fieldToDisplay: fieldData });
  }

  addFieldHandler = (groupName: string, somefunction: Function) => {
    console.log({ groupName });
    this.setState({ isCreateFieldPanelOpen: true, createFieldGroupName: groupName, fieldCreationFunction: somefunction });
  }

  closeFieldCreatePanel = (fieldData: ISPField) => {
    let currentItems = this.state.ListOfGroups;
    fieldData.JustAdded = true;
    let isAsc = true;
    let groupDoesNotExist: boolean = currentItems.find(n => n.Name == fieldData.Group) == undefined ? true : false;
    if (groupDoesNotExist) {
      currentItems.push({ Name: fieldData.Group, Fields: [fieldData as ISPField], Ascending: true } as IGroup);
      this.setState({ isCreateFieldPanelOpen: false, ListOfGroups: currentItems });
    }
    else {
      let index = currentItems.findIndex(n=>n.Name == fieldData.Group);
      currentItems[index].Fields = [...currentItems[index].Fields, fieldData];
      isAsc = currentItems[index].Ascending;
      this.sortGroupFields(fieldData.Group, isAsc);
      this.state.fieldCreationFunction(fieldData);
      this.setState({ isCreateFieldPanelOpen: false, ListOfGroups: currentItems });
    }
  }

  closePanel = () => {
    this.setState({ isCreateFieldPanelOpen: false });
  }

  sortGroupFields = (groupName, ascending: boolean) => {
    function compare(a, b) {
      const val1 = a.Title.toUpperCase();
      const val2 = b.Title.toUpperCase();

      let comparison = 0;
      if (val1 > val2) {
        comparison = ascending ? 1 : -1;
      } else if (val1 < val2) {
        comparison = ascending ? -1 : 1;
      }
      return comparison;
    }
    let currentItems = this.state.ListOfGroups;
    currentItems.forEach((item) => {
      if (item.Name == groupName) {
        item.Fields.sort(compare);
        item.Ascending = ascending;
      }
    });
    this.setState({ ListOfGroups: currentItems });
  }

  protected async deleteField(id, groupName): Promise<any> {
    console.log('delete field ', id);
    let context = this.props.context;

    const h2 = new Headers();
    h2.append("Accept", "application/json;odata.metadata=full");
    h2.append("Content-type", "application/json;odata.metadata=full");
    h2.append("IF-MATCH", "*");
    h2.append("X-HTTP-Method", "DELETE");

    const optUpdate1: ISPHttpClientOptions = {
      headers: h2
    };
    let response = await context.spHttpClient.post(context.pageContext.web.absoluteUrl + `/_api/web/fields/getbyid(guid'${id}')`, SPHttpClient.configurations.v1, optUpdate1);

    if (response.status == 204) {
      console.log("Field deleted ", id);
      let currentItems = this.state.ListOfGroups;
      currentItems.forEach((item) => {
        if (item.Name == groupName) {
          item.Fields = item.Fields.filter(n => n.Id != id);
        }
      });
      return true;
    }
    return false;
  }

  public render(): React.ReactElement<IFieldManagementProps> {
    let { groups, items } = this.state.fieldsPlain;///this.magicWithGroups(this.mockData);

    return (
      <div className={styles.fieldManagement}>
        <Panel headerText="Create new site column" isOpen={this.state.isCreateFieldPanelOpen} type={PanelType.medium} onDismiss={() => this.closePanel()}>
          <FieldCreate context={this.props.context} group={this.state.createFieldGroupName} onItemSaved={this.closeFieldCreatePanel} closePanel={this.closePanel} />
        </Panel>
        <Panel isOpen={this.state.isPanelOpened} type={PanelType.medium} onDismiss={() => this.setState({ isPanelOpened: false })}>
          <FieldDisplay field={this.state.fieldToDisplay} />
        </Panel>
        {
          this.state.ListOfGroups.length == 0 ? 
          <div>
            <Spinner label="Data is being loaded..." />
          </div>
          :
          this.state.ListOfGroups.map((group) => {
            return (<DisplayGroups group={group} key={group.Name} addFieldHandler={this.addFieldHandler} removeFieldHandler={this.deleteField.bind(this)} />);
          })
        }
      </div>
    );
  }

  protected groupBy = key => array =>
    array.reduce((objectsByKeyValue, obj) => {
      const value = obj[key];
      objectsByKeyValue[value] = (objectsByKeyValue[value] || []).concat(obj);
      return objectsByKeyValue;
    }, {})

  protected async _retrieveColumns(): Promise<any> {
    let context = this.props.context;
    let requestUrl = context.pageContext.web.absoluteUrl + `/_api/web/fields?&$orderBy=Title`; //?$filter=CanBeDeleted eq true`;

    let response: SPHttpClientResponse = await context.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1);

    function compare(a, b) {
      const val1 = a.Name.toUpperCase();
      const val2 = b.Name.toUpperCase();

      let comparison = 0;
      if (val1 > val2) {
        comparison = 1;
      } else if (val1 < val2) {
        comparison = -1;
      }
      return comparison;
    }

    if (response.ok) {
      let responseJSON = await response.json();
      if (responseJSON != null && responseJSON.value != null) {
        const refineObjectByGroups = this.groupBy('Group');
        let refinedGroups = refineObjectByGroups(responseJSON.value);
        let groupsArray: IGroup[] = [];
        for (let key in refinedGroups) {
          if (key == "_Hidden") continue;
          let obj = refinedGroups[key];
          groupsArray.push({ Name: key, Ascending: true, Fields: (obj as ISPField[]) });
        }
        groupsArray.sort(compare);
        this.setState({ ListOfGroups: groupsArray });
      }
    }
  }
}