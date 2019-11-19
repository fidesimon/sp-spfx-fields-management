import * as React from 'react';
import styles from './FieldManagement.module.scss';
import { IFieldManagementProps } from './IFieldManagementProps';
import { Panel, PanelType } from 'office-ui-fabric-react';
import { IGroup } from './Group';
import { GroupList } from './GroupList';
import FieldDisplay from './FieldDisplay';

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
}

export default class FieldManagement extends React.Component<IFieldManagementProps, IFieldManagementState> {
  constructor(props){
    super(props);
    this.state = { ListOfGroups: [], isPanelOpened: false, fieldToDisplay: null}
  }

  componentWillMount(){
    const mockData : Array<IGroup> = [
      { Name: "asdf", Fields: [{Title: "some name", Id: "some id"}] },
      { Name: "qwert", Fields: [{Title: "some name", Id: "some id"}]},
    ];
    this.setState({ ListOfGroups: mockData });
    const grupki = this._retrieveColumns();
  }

  handleFieldClick = (fieldData : ISPField) => {
    console.log({fieldData});
    this.setState({isPanelOpened: true, fieldToDisplay: fieldData});
  }

  public render(): React.ReactElement<IFieldManagementProps> {
    

    return (
      <div className={ styles.fieldManagement }>
        <Panel isOpen={this.state.isPanelOpened} type={PanelType.medium} onDismiss={() => this.setState({isPanelOpened: false})}>
          <FieldDisplay field={this.state.fieldToDisplay} />
        </Panel>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <GroupList groups={this.state.ListOfGroups} clickHandler={this.handleFieldClick} />
            </div>
          </div>
        </div>
      </div>
    );
  }

  protected groupBy = key => array =>
    array.reduce((objectsByKeyValue, obj) =>{
      const value = obj[key];
      objectsByKeyValue[value] = (objectsByKeyValue[value] || []).concat(obj);
      return objectsByKeyValue;
    }, {});

  protected async _retrieveColumns(): Promise<any> {
    let context = this.props.context;
    let requestUrl = context.pageContext.web.absoluteUrl + `/_api/web/fields`; //?$filter=CanBeDeleted eq true`;

    let response : SPHttpClientResponse = await context.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1);
    
    if(response.ok){
      let responseJSON = await response.json();
      if(responseJSON != null && responseJSON.value != null){
        const refineObjectByGroups = this.groupBy('Group');
        let refinedGroups = refineObjectByGroups(responseJSON.value);
        let groupsArray : IGroup[] = [];
        for (let key in refinedGroups){
          if(key == "_Hidden") continue;
          let obj = refinedGroups[key];
          groupsArray.push({Name: key, Fields: (obj as ISPField[])});
        }
        this.setState({ListOfGroups: groupsArray});
      }
    }
  }
}
