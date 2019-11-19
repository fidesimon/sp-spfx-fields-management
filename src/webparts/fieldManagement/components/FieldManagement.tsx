import * as React from 'react';
import styles from './FieldManagement.module.scss';
import { IFieldManagementProps } from './IFieldManagementProps';
import { Panel } from 'office-ui-fabric-react';
import { IGroup } from './Group';
import { GroupList } from './GroupList';

import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions
} from '@microsoft/sp-http';
import { ISPField } from './SPField';

export interface IFieldManagementState {
  ListOfGroups: IGroup[];
  isPanelOpened: boolean;
}

export default class FieldManagement extends React.Component<IFieldManagementProps, IFieldManagementState> {
  constructor(props){
    super(props);
    this.state = { ListOfGroups: [], isPanelOpened: false}
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
    this.setState({isPanelOpened: true});
  }

  public render(): React.ReactElement<IFieldManagementProps> {
    

    return (
      <div className={ styles.fieldManagement }>
        <Panel isOpen={this.state.isPanelOpened} onDismiss={() => this.setState({isPanelOpened: false})} />
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
