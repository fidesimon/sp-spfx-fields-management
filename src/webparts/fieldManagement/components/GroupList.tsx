import * as React from 'react';
import { IGroup, Group } from './Group';

export interface GroupListProps{
    groups: IGroup[];
    clickHandler: Function;
  }
  
  
  
  export class GroupList extends React.Component<GroupListProps, {}> {
    
    public render(): React.ReactElement<GroupListProps> {
      const groups = this.props.groups;
      return(
        <div>
        {groups.map(group => <Group key={group.Name} name={group.Name} fields={group.Fields} clickHandler={this.props.clickHandler} />)}
        </div>
      );
    }
  }