import * as React from 'react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';

export interface FieldCreateProps {
    group: string,
}

export interface FieldCreateState{
    columnName: string,
    internalName: string,
    group: string,
    description: string,
    required: boolean,
    enforceUniqueValues: boolean,
    maxNoCharacters: number,
    defaultValue: string
}

enum asd{
    columnName,
    internalName,       
}

export default class FieldCreate extends React.Component<FieldCreateProps, FieldCreateState>{
    constructor(props){
        super(props);
        this.state = { columnName: '', internalName: '', group: this.props.group, description: '', required: false, enforceUniqueValues: false, maxNoCharacters: 255, defaultValue: ''}
    }



    generateInternalName(){
        let columnName: string = (document.getElementById("columnName") as HTMLInputElement).value;
        let newValue: string =  columnName.replace(/[^A-Z0-9]+/ig, "");
        this.setState({columnName: columnName, internalName: (newValue.length >= 32 ? newValue.substr(0, 32) : newValue)});
        //Add column internalName validation check. + possible counter
    }

    render() {
        return (
            <div>
                <span>Create site column</span>
                <TextField label="Column Name" id="columnName" required value={this.state.columnName} onKeyUp={() => this.generateInternalName()} />
                <TextField label="Internal Name" required value={this.state.internalName} onKeyUp={(evt) => this.setState({internalName: (evt.target as HTMLInputElement).value})} />
                <TextField label="Group" defaultValue={this.props.group} onKeyUp={(evt) => this.setState({group: (evt.target as HTMLInputElement).value})} />
                <TextField label="Description" name="columnName" multiline autoAdjustHeight onKeyUp={(evt) => this.setState({description: (evt.target as HTMLInputElement).value})} />
                <Toggle label="Required" onChanged={(evt) => this.setState({required: evt})} />
                <Toggle label="Enforce Unique Values?" onChanged={(evt) => this.setState({enforceUniqueValues: evt})} />
                <TextField label="Maximum number of characters" type="number" defaultValue="255" onKeyUp={(evt) => this.setState({maxNoCharacters: (evt.target as HTMLInputElement).valueAsNumber})} />
                <TextField label="Default value" onChange={(evt) => console.log({evt})} />
            </div>
        );
    }
}

/*
For TextField options in SharePoint:
1. Column Name
2. Group - existing group drop-down or new group
3. Description
4. Required
5. Enforce unique values
6. Maximum number of characters (def 255) 
7. Default Value - text or calculated value - calculated value skipped until v2
8. Column formatting (json) - skipped until v2
9. Column validation: Formula and User Message.   - skipped until v2

FieldTypeKind: 2


*/