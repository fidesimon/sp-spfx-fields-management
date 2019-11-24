import * as React from 'react';
import SPField, { ISPField } from './SPField';
import styles from './FieldManagement.module.scss';
import { DetailsList } from 'office-ui-fabric-react';

export interface IGroup{
    Name: string;
    Fields?: ISPField[];
    InternalName?: string;
    GroupId?: string;
  }
  
  export type GroupProps = {
    name: string
    fields?: ISPField[],
    clickHandler: Function,
    addFieldHandler: Function
  }

  export class Group extends React.Component<GroupProps, {}> {
    public render(): React.ReactElement<GroupProps> {
      const groupName = this.props.name;
      const fields = this.props.fields;
      return(
        <div className={styles.container}>
          <div className={styles.groupHeader}>
            <div className={styles.groupName}>{groupName}</div>
            <div onClick={()=>this.props.addFieldHandler(groupName)} className={styles.pullRight}>Add New Field</div>
          </div>
          { fields.map(field => <SPField key={field.Id} field={field} clickHandler={this.props.clickHandler} />)}
        </div>
      );
    }
  }