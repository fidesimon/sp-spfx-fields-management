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
  fieldsAscending: boolean,
  deleteField: Function,
  groupExpanded: boolean
};
  
export const Group: React.FC<GroupProps> = (props) => {
  const [fieldsArray, setFieldsArray] = React.useState(props.fields);
  const [ascendingSort, setAscendingSort] = React.useState(props.fieldsAscending);
  const [groupExpanded, setGroupExpanded] = React.useState(props.groupExpanded);

  const sortGroupFields = (ascending: boolean) => {
    function compare(a: ISPField, b: ISPField) {
      const val1 = a.Title.toUpperCase();
      const val2 = b.Title.toUpperCase();
    
      let comparison = 0;
      if (val1 > val2)
        comparison = ascending ? 1 : -1;
      else if (val1 < val2)
        comparison = ascending ? -1 : 1;
      return comparison;
    }

    setFieldsArray(fieldsArray.sort(compare));
    setAscendingSort(!ascendingSort);
  }

  const groupName = props.name;
  const fields = props.fields;
  return(
    <div className={styles.container}>
      <div className={styles.groupHeader}>
        <div className={styles.groupName}>
          <div className={styles.sort} onClick={()=>sortGroupFields(!ascendingSort)}>
            {ascendingSort ? 
              <Icon iconName="Ascending" className={styles.sortingIcon} />
              :
              <Icon iconName="Descending" className={styles.sortingIcon} />
            }
            <span onClick={()=>setGroupExpanded(!groupExpanded)}>
            {
              `${groupName} (${props.fields.length})`
            }
            </span>
          </div>
          
        </div>
        
        <div onClick={()=>props.addFieldHandler(groupName)} className={styles.pullRight}>Add New Field</div>
      </div>

      { 
        groupExpanded ?
          fields.map(field => <SPField key={field.Id} field={field} clickHandler={props.clickHandler} deleteField={props.deleteField} />)
        : null
      }
    </div>
  );
}