import * as React from 'react';
import { IGroup, Group } from './Group';

export interface GroupListProps{
    groups: IGroup[];
    clickHandler: Function;
    addFieldHandler: Function;
  }
  
  
  
  export class GroupList extends React.Component<GroupListProps, {}> {
    
    public render(): React.ReactElement<GroupListProps> {
      const groups = this.props.groups;
      return(
        <div>placeholder..</div>
      );
      /*return(
        <div>
        {groups.map(group => <Group key={group.Name} name={group.Name} fields={group.Fields} addFieldHandler={this.props.addFieldHandler} clickHandler={this.props.clickHandler} />)}
        </div>
      );*/
    }
  }