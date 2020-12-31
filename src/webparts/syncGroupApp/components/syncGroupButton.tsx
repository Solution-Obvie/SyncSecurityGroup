import * as React from 'react';
import styles from './SyncGroupApp.module.scss';
import { disableBodyScroll, PrimaryButton } from 'office-ui-fabric-react';
import { HttpClient, SPHttpClient, HttpClientConfiguration, HttpClientResponse, ODataVersion, IHttpClientConfiguration, IHttpClientOptions, ISPHttpClientOptions } from '@microsoft/sp-http'; 
import { TooltipHost, ITooltipHostStyles, DirectionalHint } from 'office-ui-fabric-react/lib/Tooltip';
import { useId } from '@uifabric/react-hooks';

export default function SyncGroupButton(props){  

  const hostStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline-block', width:'100%' } };
  const tooltipId = useId('tooltip');

    var functionUrl = "https://powershellgroupoperation.azurewebsites.net/api/CompareGroup";    
    function callAzureFunction() {    
          const requestHeaders: Headers = new Headers();    
          requestHeaders.append("Content-type", "application/json");  
          requestHeaders.append("Cache-Control", "no-cache");    
            
          var siteUrl: string = props.context.pageContext.web.absoluteUrl;      
          var body = `{ microsoftgroupID:  '${props.group.ID}', securitygroupID:  '${props.group.SecurityGroupID}', siteUrl: '${siteUrl}'}`
          console.log(body)
            const postOptions: IHttpClientOptions = {    
            headers: requestHeaders,
            body:`{ microsoftgroupID:  '${props.group.ID}', securitygroupID:  '${props.group.SecurityGroupID}', siteUrl: '${siteUrl}'}`
          };    
            
            props.context.httpClient.post(functionUrl, HttpClient.configurations.v1, postOptions).then((response) =>{     
             console.log(response) 
             console.log(response.nativeResponse.status) 
             props.setProgress(false) 
            })    
            
                .catch ((response: any) => {    
                let errMsg: string = `WARNING - error when calling URL ${functionUrl}. Error = ${response.message}`;    
                console.log(errMsg)
              });       
      }    


   

    function SyncGroup(){
      callAzureFunction();
      props.setProgress(true)
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
    
</div>
  

    
)

}