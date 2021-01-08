import * as React from 'react';
import styles from './SyncGroupApp.module.scss';
import { sp } from "@pnp/sp";
import { Item } from '@pnp/sp/items';

function GroupInformation(props){




    React.useEffect(() =>{
        getItems();
    }, [])



    function getItems() {

        sp.web.select("AllProperties").expand("AllProperties").get().then(function(result){  
            // Select the AllProperties from the result
            console.log(result["AllProperties"]);
            console.log(result["AllProperties"].MicrosoftGroup)
            var MicrosoftGroup = JSON.parse(result["AllProperties"].MicrosoftGroup)
            var SecurityGroup = JSON.parse(result["AllProperties"].SecurityGroupLinked)
            var LastSync = result["AllProperties"].LastSync
            
            props.setGroup({"Title": MicrosoftGroup.Name, "ID": MicrosoftGroup.Id, 
            "SecurityGroupTitle":SecurityGroup.Name,"SecurityGroupID":SecurityGroup.Id, "LastSync":LastSync}); 
            //console.log(props.group)
        }); 

    }

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

export default GroupInformation;