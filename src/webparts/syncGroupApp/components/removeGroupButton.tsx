import * as React from 'react';
import styles from './SyncGroupApp.module.scss';
import { PrimaryButton } from 'office-ui-fabric-react';
import { sp } from "@pnp/sp";
import { HttpClient, SPHttpClient, HttpClientConfiguration, HttpClientResponse, ODataVersion, IHttpClientConfiguration, IHttpClientOptions, ISPHttpClientOptions } from '@microsoft/sp-http'; 

export default function RemoveGroupButton(props){  


    function getItems() {
        sp.web.lists.getByTitle("syncGroupAppSettings").items.get().then(items => {
            items.forEach(item => {
                console.log(item)
               props.setGroup({"Title": item.Title, "ID": item.MicrosoftGroupID, "isSecurityGroup": item.isSecurityGroup,
               "SecurityGroupTitle":item.SecurityGroupName,"SecurityGroupID":item.SecurityGroupID});
            });    
            props.setProgress(false)    
        })
    }

    var functionUrl = "https://powershellgroupoperation.azurewebsites.net/api/RemoveSecurityGroup?code=7pN0k7aOKTH2Kg9hkj12zuiKe9kvKRXfiTLQb1rdCa7Tojcvw4E9Nw==";    
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
             if(response.nativeResponse.status == 200)
             {
                 getItems();
             }
            })    
            
                .catch ((response: any) => {    
                let errMsg: string = `WARNING - error when calling URL ${functionUrl}. Error = ${response.message}`;    
                console.log(errMsg)
              });       
      }    


    function RemoveGroup(){
        props.setProgress(true)
      callAzureFunction();
    }
    
    return(
    
        <PrimaryButton className={styles.removeButton} text="Remove" onClick={RemoveGroup} disabled={props.progress}/>
    )
    
    }