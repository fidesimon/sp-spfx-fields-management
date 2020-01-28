import * as React from 'react';
import styles from './GroupHeader.module.scss';
import { Separator, IGroup } from 'office-ui-fabric-react';

interface IGroupHeaderProps {
    groupName: string;
}

export default class GroupHeader extends React.Component<IGroupHeaderProps,{}>{
    constructor(props: IGroupHeaderProps) {
        super(props);
        
    }

    render(){
        return (
            <div className={styles.groupHeader}>
                <div className={styles.container}>
                    <div className={styles.row}>
                        {this.props.groupName}
                    </div>
                </div>
            </div>
        );
    }
}