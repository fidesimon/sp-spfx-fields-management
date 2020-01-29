import * as React from 'react'
import { CommandBar } from 'office-ui-fabric-react';
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
                <CommandBar
                    items={
                        [{
                            key: 'newItem',
                            text: 'New',
                            cacheKey: 'myCacheKey', // changing this key will invalidate this item's cache
                            iconProps: { iconName: 'Add' },
                            onClick: () => this.props.addFieldHandler(this.props.group.Name)
                        },
                        {
                            key: 'delete',
                            text: 'Delete',
                            iconProps: { iconName: 'Delete' },
                            disabled: true
                        }]
                    }
                    ariaLabel="Use left and right arrow keys to navigate between commands"
                />
                <DisplayFields fields={this.props.group} removeFieldHandler={this.props.removeFieldHandler} />
            </>
        );
    }
}