import { groupBy } from '@microsoft/sp-lodash-subset';
import * as React from 'react';
import styles from './SyncGroupApp.module.scss';

export default function LastSync(props){ 

return(
    
    <div className={styles.flexAround100}>
        <div className={styles.flexColumn30}>
        <div className={styles.fontWeigth600}>User(s) Added</div>
        {
         props.group.AddedMembers.map((user) =>      
            <div className={styles.marginTop5}> {user}</div>
        )

        }
        </div>

        <div className={styles.flexColumn30}>
        <div className={styles.fontWeigth600}>User(s) Removed</div>
        {
         props.group.RemovedMembers.map((user) =>      
            <div className={styles.marginTop5}> {user}</div>
        )

        }
        </div>

    </div>
    
)

}