import * as React from 'react';
import styles from './SyncGroupApp.module.scss';
import { ISyncGroupAppProps } from './ISyncGroupAppProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { MSGraphClient } from '@microsoft/sp-http';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import { ISiteUser } from '@pnp/sp/site-users';
import SelectSecurity from './selectSecurity'
import GroupInformation from './groupInformation'
import SyncGroupAppWebPart from '../SyncGroupAppWebPart';
import SyncGroupButton from './syncGroupButton'
import RemoveGroupButton from './removeGroupButton'
import SecurityGroupInformation from './securityGroupInformation'
import { ProgressIndicator } from 'office-ui-fabric-react/lib/ProgressIndicator';


export default function SyncGroupApp(props){  

  const [group, setGroup] = React.useState({"Title" : "", 
  "ID" : "",
  "isSecurityGroup": false,
  "SecurityGroupTitle":"",
  "SecurityGroupID":""});

  const [progress, setProgress] = React.useState(false);
    
    return (
      <div className={ styles.syncGroupApp }>
        <div className={ styles.container }>
         <div className={styles.titleContainer}>
           <div className={styles.appTitle}>
           Synchronise your group application 
           </div>
           {/* <div className={styles.underlineTitle}></div> */}
         </div>

        <div className={styles.actionContainer}>
        {
           !group.isSecurityGroup && 
           <div>
                <GroupInformation group={group} setGroup={setGroup}/> 
               <SelectSecurity context={props.context} group={group} setGroup={setGroup}  setProgress={setProgress} progress={progress}/>   
            </div>  
         }  
         {
         group.isSecurityGroup && 
         <div className={styles.flexAround}>
           <div className={styles.flexColumn}>
           <GroupInformation group={group} setGroup={setGroup}/> 
           <SyncGroupButton className={styles.width50} context={props.context} group={group} setProgress={setProgress} progress={progress}/>
            
            </div>
         <div className={styles.flexColumn}>
            <SecurityGroupInformation group={group}/>
           <RemoveGroupButton className={styles.width50} context={props.context} group={group} setGroup={setGroup} setProgress={setProgress} progress={progress}/>
         </div> 
         </div>
         }
       
       
         {
           progress &&
           <ProgressIndicator label="Sending Request.." description="A request has been sent. Please don't leave the page." />
         }
        </div>

      

        </div>
      </div>
    );

}
