import * as React from 'react';
import styles from './SyncGroupApp.module.scss';
import { sp } from "@pnp/sp";
import { Item } from '@pnp/sp/items';

function GroupInformation(props){

  


    React.useEffect(() =>{
        getItems();
    }, [])



    function getItems() {
        sp.web.lists.getByTitle("syncGroupAppSettings").items.get().then(items => {
            items.forEach(item => {
                console.log(item)
               props.setGroup({"Title": item.Title, "ID": item.MicrosoftGroupID, "isSecurityGroup": item.isSecurityGroup,
               "SecurityGroupTitle":item.SecurityGroupName,"SecurityGroupID":item.SecurityGroupID});
            });        
        })
    }

    return(
        <div>
            <div>Name of your group :</div>
            <div>{props.group.Title}</div>
        </div>
    )

}

export default GroupInformation;