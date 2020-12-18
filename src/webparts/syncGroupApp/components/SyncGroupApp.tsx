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

export default function SyncGroupApp(props){  

  const [group, setGroup] = React.useState({"Title" : "", 
  "ID" : "",
  "isSecurityGroup": false,
  "SecurityGroupTitle":"",
  "SecurityGroupID":""});

    
    return (
      <div className={ styles.syncGroupApp }>
        <div className={ styles.container }>
         <div className={styles.appTitle}>
           Synchronise your group application
         </div>
         <div>

         </div>
         {
           !group.isSecurityGroup && 
           <SelectSecurity context={props.context} group={group}/>
         } 
         <GroupInformation group={group} setGroup={setGroup}/>
        </div>
      </div>
    );

}
