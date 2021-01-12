import * as React from 'react';
import { sp } from "@pnp/sp";
import styles from './SyncGroupApp.module.scss';
import { disableBodyScroll, PrimaryButton } from 'office-ui-fabric-react';
import { HttpClient, SPHttpClient, HttpClientConfiguration, HttpClientResponse, ODataVersion, IHttpClientConfiguration, IHttpClientOptions, ISPHttpClientOptions } from '@microsoft/sp-http'; 
import { TooltipHost, ITooltipHostStyles, DirectionalHint } from 'office-ui-fabric-react/lib/Tooltip';
import { useId } from '@uifabric/react-hooks';
import { getItem } from './functions/updateItem'
import {callAzureFunction} from './functions/callAzureFunction'

export default function SyncGroupButton(props){  

  const hostStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline-block', width:'100%' } };
  const tooltipId = useId('tooltip');
    var functionUri = "https://powershellgroupoperation.azurewebsites.net/api/CompareGroup";     
    function SyncGroup(){
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
        content="This button will send a request to synchronize your group with the security group linked. This will remove members which are no longer in the security group and add new that were added to the security group."
        // This id is used on the tooltip itself, not the host
        // (so an element with this id only exists when the tooltip is shown)
        id={tooltipId}
        // calloutProps={calloutProps}
         styles={hostStyles}
         directionalHint={DirectionalHint.bottomCenter}
         
      >
          <PrimaryButton aria-describedby={tooltipId} className={styles.syncButton} text="Sync" onClick={SyncGroup} disabled={props.progress}/>
      </TooltipHost>
      <div>
       
      </div>
    
</div>
  

    
)

}