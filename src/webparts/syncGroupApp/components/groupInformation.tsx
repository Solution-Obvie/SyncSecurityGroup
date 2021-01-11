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
            var AddedMembers = result["AllProperties"].AddedMember
            if(AddedMembers != " "){
                AddedMembers = []
               // AddedMembers.push(JSON.parse(result["AllProperties"].AddedMember))
               if(Array.isArray(JSON.parse(result["AllProperties"].AddedMember))){
                JSON.parse(result["AllProperties"].AddedMember).forEach(element => {
                    AddedMembers.push(element)
                });
               }
               else{
                AddedMembers.push(JSON.parse(result["AllProperties"].AddedMember))
               }
              
            }
            else{
                AddedMembers = []
            }
            var RemovedMembers = result["AllProperties"].RemovedMember
            if(RemovedMembers != " "){
                RemovedMembers = []
           
                if(Array.isArray(JSON.parse(result["AllProperties"].RemovedMember))){
                    JSON.parse(result["AllProperties"].RemovedMember).forEach(element => {
                        RemovedMembers.push(element)
                    });
                   }
                   else{
                    RemovedMembers.push(JSON.parse(result["AllProperties"].RemovedMember))
                   }
                //RemovedMembers.push(JSON.parse(result["AllProperties"].RemovedMember))
            }
            else{
                RemovedMembers = []
            }

            
            
            props.setGroup({"Title": MicrosoftGroup.Name, "ID": MicrosoftGroup.Id, 
            "SecurityGroupTitle":SecurityGroup.Name,"SecurityGroupID":SecurityGroup.Id, "LastSync":LastSync,
            "AddedMembers":  AddedMembers, "RemovedMembers" : RemovedMembers  }); 
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