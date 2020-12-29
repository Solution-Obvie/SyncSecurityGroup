import * as React from 'react';
import styles from './SyncGroupApp.module.scss';
import { disableBodyScroll, PrimaryButton } from 'office-ui-fabric-react';
import { HttpClient, SPHttpClient, HttpClientConfiguration, HttpClientResponse, ODataVersion, IHttpClientConfiguration, IHttpClientOptions, ISPHttpClientOptions } from '@microsoft/sp-http'; 

export default function SyncGroupButton(props){  

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

    <PrimaryButton className={styles.syncButton} text="Sync" onClick={SyncGroup} disabled={props.progress}/>
)

}