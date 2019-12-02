import * as React from 'react';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import styles from './FieldManagement.module.scss';

export type SPFieldProps = {
    field: ISPField;
    clickHandler: Function;
    deleteField: Function;
  }

export default class SPField extends React.Component<SPFieldProps, {}>{
    constructor(props){
      super(props);
    }
  
    public render(): React.ReactElement<SPFieldProps>{
      const field = this.props.field;
      return(
        <div className="ms-Grid-row" onClick={() => this.props.clickHandler(field)} style={{backgroundColor: this.props.field.JustAdded ? 'limegreen' : 'white'}}>
          <div className="ms-Grid-col ms-lg6">{field.Title}</div>
          <div className="ms-Grid-col ms-lg5">{field.TypeDisplayName}</div>
          <div className="ms-Grid-col ms-lg1">
            {
              this.props.field.CanBeDeleted ?
              <Icon iconName="Delete" style={{color: '#ff0000'}} onClick={(ev) => {
                ev.stopPropagation();
                this.props.deleteField(this.props.field.Id, this.props.field.Group)} 
              } />
              :
              <Icon iconName="Delete" style={{color: '#d8d8d8'}} />
            }
          </div>
        </div>
      );
    }
  }
  
  export interface ISPField {
    JustAdded?: boolean;
    '@odata.type'?: string;
    '@odata.id'?: string;
    '@odata.editLink'?: string;
    UnlimitedLengthInDocumentLibrary?: boolean;
    AllowHyperlink?: boolean;
    AppendOnly?: boolean;
    NumberOfLines?: number;
    RichText?: boolean;
    AutoIndexed?: boolean; 
    CanBeDeleted?: boolean;
    Choices?: string[];
    ClientSideComponentId?: string;
    ClientSideComponentProperties?: string;
    ClientValidationFormula?: string;
    ClientValidationMessage?: string;
    CustomFormatter?: string;
    DefaultFormula?: string;
    DefaultValue?: string;
    Description?: string;
    Direction?: string;
    EnforceUniqueValues?: boolean;
    EntityPropertyName?: string;
    FieldTypeKind?: number;
    Filterable?: boolean;
    FromBaseType?: boolean;
    Group?: string;
    Hidden?: boolean;
    Id?: string;
    IndexStatus?: number;
    Indexed?: boolean;
    InternalName?: string;
    JSLink?: string;
    MaxLength?: number;
    PinnedToFiltersPane?: boolean;
    ReadOnlyField?: boolean;
    Required?: boolean;
    SchemaXml?: string;
    Scope?: string;
    Sealed?: boolean;
    ShowInFiltersPane?: number;
    Sortable?: boolean;
    StaticName?: string;
    Title?: string;
    TypeAsString?: string;
    TypeDisplayName?: string;
    TypeShortDescription?: string;
    ValidationFormula?: string;
    ValidationMessage?: string;
    DisplayFormat?: number;
    MaximumValue?: number;
    MinimumValue?: number;
    ShowAsPercentage?: boolean;
  }