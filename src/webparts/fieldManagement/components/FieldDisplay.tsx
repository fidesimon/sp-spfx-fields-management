import * as React from 'react';
import SPField, { ISPField } from './SPField';

export type FieldDisplayProps = {
    field: ISPField;
}

export default class FieldDisplay extends React.Component<FieldDisplayProps, {}>{
    render(){
        const field = this.props.field;
        return (
            <div className="ms-Grid" dir="ltr">
                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg2">Field Name:</div>
                    <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg10">{field.Title}</div>
                </div>
                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg2">Internal Name:</div>
                    <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg10">{field.InternalName}</div>
                </div>
                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg2">Field Type:</div>
                    <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg10">{field.TypeDisplayName}</div>
                </div>
                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg2">Field Id:</div>
                    <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg10">{field.Id}</div>
                </div>
                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg2">Group:</div>
                    <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg10">{field.Group}</div>
                </div>
            </div>
        );
    }
}