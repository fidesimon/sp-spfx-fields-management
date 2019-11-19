import * as React from 'react';

export type SPFieldProps = {
    field: ISPField;
    clickHandler: Function;
  }

export default class SPField extends React.Component<SPFieldProps, {}>{
    constructor(props){
      super(props);
    }
  
    public render(): React.ReactElement<SPFieldProps>{
      const field = this.props.field;
      return(
        <div>
          <div style={{"paddingLeft": '20px'}} onClick={() => this.props.clickHandler(field)}>
            {field.Title} <small>({field.StaticName})XXXXnpm </small>
          </div>
        </div>
      );
    }
  }
  

  
  export type ISPField = {
    '@odata.type'?: string;
    '@odata.id'?: string;
    '@odata.editLink'?: string;
    AutoIndexed?: boolean; 
    CanBeDeleted?: boolean;
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
  }