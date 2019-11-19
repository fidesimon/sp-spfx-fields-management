import * as React from 'react';
import SPField, { ISPField } from './SPField';

export type FieldDisplayProps = {
    field: ISPField;
}

export default class FieldDisplay extends React.Component<FieldDisplayProps, {}>{
    render(){
        const field = this.props.field;
        return (
            <div>
                Field Name: <b>{field.Title}</b><br />
                Internal Name: <b>{field.InternalName}</b><br />
                Field Type: <b>{field.TypeDisplayName}</b><br />
                Id: <b>{field.Id}</b><br />
                Group: <b>{field.Group}</b><br />
            </div>

        );
    }
}