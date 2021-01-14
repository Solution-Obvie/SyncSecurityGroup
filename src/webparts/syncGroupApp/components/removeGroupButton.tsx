import * as React from 'react';
import styles from './SyncGroupApp.module.scss';
import { PrimaryButton } from 'office-ui-fabric-react';
import { sp } from "@pnp/sp";
import { HttpClient, SPHttpClient, HttpClientConfiguration, HttpClientResponse, ODataVersion, IHttpClientConfiguration, IHttpClientOptions, ISPHttpClientOptions } from '@microsoft/sp-http'; 
import { TooltipHost, ITooltipHostStyles, DirectionalHint } from 'office-ui-fabric-react/lib/Tooltip';
import { useId } from '@uifabric/react-hooks';
import { getItem } from './functions/updateItem'
import {callAzureFunction} from './functions/callAzureFunction'

export default function RemoveGroupButton(props){  

  const hostStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline-block', width:'100%' } };
  const tooltipId = useId('tooltip2');

    //var functionUri = "https://powershellgroupoperation.azurewebsites.net/api/RemoveSecurityGroup?code=7pN0k7aOKTH2Kg9hkj12zuiKe9kvKRXfiTLQb1rdCa7Tojcvw4E9Nw==";  
   
    var functionUri = "https://powershellgroupoperation.azurewebsites.net/api/RemoveSecurityGroup"
    
    function RemoveGroup(){
      props.setProgress(true)
      callAzureFunction(functionUri, props.context, props.group.ID, props.group.SecurityGroupID)
      .then(data =>     
        getItem()
        .then(data => {
          props.setGroup(data)
          props.setProgress(false)
        }
        )
      )  
    }
    
    return( 
    
      <div className={styles.width50}>
        <TooltipHost
        className={styles.width100}
        content="This button will send a request to remove the security group from your group. This will remove all members that belong to the security group, but not members that were inside your group before the link."
        // This id is used on the tooltip itself, not the host
        // (so an element with this id only exists when the tooltip is shown)
        id={tooltipId}
        // calloutProps={calloutProps}
         styles={hostStyles}
         directionalHint={DirectionalHint.bottomCenter}
        >
        <PrimaryButton aria-describedby={tooltipId} className={styles.removeButton} text="Remove" onClick={RemoveGroup} disabled={props.progress}/>
        </TooltipHost>
        </div>
    )
    
    }