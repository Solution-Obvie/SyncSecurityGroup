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
           <div className={styles.underlineTitle}></div>
         </div>

        

         {
           !group.isSecurityGroup && 
           <SelectSecurity context={props.context} group={group} setGroup={setGroup}  setProgress={setProgress}/>
         } 
         <GroupInformation group={group} setGroup={setGroup}/>
         {
         group.isSecurityGroup && 
         <div className={styles.buttonContainer}>
           <SyncGroupButton context={props.context} group={group}/>
           <RemoveGroupButton context={props.context} group={group} setGroup={setGroup} setProgress={setProgress}/>
         </div>
         }
         {
           progress &&
           <ProgressIndicator label="Sending Request.." description="A request has been sent. Please don't leave the page." />
         }

        </div>
      </div>
    );

}
