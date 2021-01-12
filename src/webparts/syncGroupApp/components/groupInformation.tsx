import * as React from 'react';
import styles from './SyncGroupApp.module.scss';
import { sp } from "@pnp/sp";
import { Item } from '@pnp/sp/items';
import { getItem } from './functions/updateItem'

export default function GroupInformation(props){

    const [riverInformation, setRiverInformation] = React.useState({});

    React.useEffect(() => {
        getItem()
        .then(data =>
          props.setGroup(data)
        );
       }, [])

    return(
        <div>  
            {
                props.group.SecurityGroupTitle != "" &&
                <div>
            <div>Group Name :</div>
            <div className={styles.groupName}>{props.group.Title}</div>
            </div>
            }
            
        </div>  
      
    )

}

