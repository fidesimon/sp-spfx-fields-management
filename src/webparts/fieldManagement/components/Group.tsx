import * as React from 'react';
import SPField, { ISPField } from './SPField';
import styles from './FieldManagement.module.scss';
import { Icon } from 'office-ui-fabric-react/lib/Icon';


export interface IGroup{
    Name: string;
    Fields?: ISPField[];
    InternalName?: string;
    GroupId?: string;
    Ascending: boolean;
  }
  
  export type GroupProps = {
    name: string
    fields?: ISPField[],
    clickHandler: Function,
    addFieldHandler: Function,
    sortHandler: Function,
    fieldsAscending: boolean,
    deleteField: Function
  };

  export class Group extends React.Component<GroupProps, {}> {
    public render(): React.ReactElement<GroupProps> {
      const groupName = this.props.name;
      const fields = this.props.fields;
      return(
        <div className={styles.container}>
          <div className={styles.groupHeader}>
            <div className={styles.groupName}>
              <div className={styles.sort} onClick={()=>this.props.sortHandler(groupName, !this.props.fieldsAscending)}>
                {this.props.fieldsAscending ? 
                  <Icon iconName="Ascending" className={styles.sortingIcon} />
                  :
                  <Icon iconName="Descending" className={styles.sortingIcon} />
                }
                {groupName}
              </div>
              
            </div>
            
            <div onClick={()=>this.props.addFieldHandler(groupName)} className={styles.pullRight}>Add New Field</div>
          </div>
          { fields.map(field => <SPField key={field.Id} field={field} clickHandler={this.props.clickHandler} deleteField={this.props.deleteField} />)}
        </div>
      );
    }
  }