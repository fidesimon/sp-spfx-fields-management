import * as React from 'react';
import { IGroup, Group } from './Group';
import { Button } from 'office-ui-fabric-react';

export interface GroupListProps{
    groups: IGroup[];
    clickHandler: Function;
    addFieldHandler: Function;
    deleteFieldHandler: Function;
  }

  export class GroupList extends React.Component<GroupListProps, {expanded: boolean}> {

    public render(): React.ReactElement<GroupListProps> {
      const groups = this.props.groups;
      return(
        <div>
          {groups.map(group => <Group 
                                  key={group.Name} 
                                  name={group.Name} 
                                  fields={group.Fields} 
                                  addFieldHandler={this.props.addFieldHandler} 
                                  deleteField={this.props.deleteFieldHandler}
                                  clickHandler={this.props.clickHandler} 
                                  fieldsAscending={group.Ascending}
                                  groupExpanded={false}
                                  />)}
        </div>
      );
    }
  }