import * as React from 'react';
import styles from './SyncGroupApp.module.scss';


export default function SecurityGroupInformation(props){



    return(
        <div>
        <div>Security Group Name :</div>
        <div className={styles.groupName}>{props.group.SecurityGroupTitle}</div>
        </div>
    )

}