import * as React from 'react'
import { CommandBar, TeachingBubbleBase } from 'office-ui-fabric-react';
import DisplayFields from './DisplayFields';
import GroupHeader from './GroupHeader';
import { IGroup } from '../Group';

interface IDisplayGroupProps {
    group: IGroup;
    addFieldHandler: Function;
    removeFieldHandler: Function;
}

export default class DisplayGroups extends React.Component<IDisplayGroupProps, {}>{
    constructor(props: IDisplayGroupProps) {
        super(props);
        
    }

    render() {
        return (
            <>
                <GroupHeader groupName={this.props.group.Name} />
                <DisplayFields fields={this.props.group} removeFieldHandler={this.props.removeFieldHandler} addFieldHandler={this.props.addFieldHandler} />
            </>
        );
    }
}