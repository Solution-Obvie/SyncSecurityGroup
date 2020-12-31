import * as React from 'react';
import styles from './SyncGroupApp.module.scss';
import { sp } from "@pnp/sp";
import { Item } from '@pnp/sp/items';
import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { HttpClient, SPHttpClient, HttpClientConfiguration, HttpClientResponse, ODataVersion, IHttpClientConfiguration, IHttpClientOptions, ISPHttpClientOptions } from '@microsoft/sp-http'; 
import { PrimaryButton } from 'office-ui-fabric-react';

export default function GroupActionAdd(props) {

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
  console.log(props.ID)

    var functionUrl = "https://powershellgroupoperation.azurewebsites.net/api/AddSecurityGroup";    
    function callAzureFunction() {    
          const requestHeaders: Headers = new Headers();    
          requestHeaders.append("Content-type", "application/json");  
          requestHeaders.append("Cache-Control", "no-cache");    
        
          var siteUrl: string = props.context.pageContext.web.absoluteUrl;    
          // var userName = "test";
          // var body = JSON.stringify({"name": "Test"})
          // console.log(body)
          //  console.log(`SiteUrl: '${siteUrl}', UserName: '${userName}'`);   
          console.log(`{ microsoftgroupID:  '${props.group.ID}', securitygroupID:  '${props.ID}', siteUrl: '${siteUrl}'}`) 
            const postOptions: IHttpClientOptions = {    
            headers: requestHeaders,
            body:`{ microsoftgroupID:  '${props.group.ID}', securitygroupID:  '${props.ID}', siteUrl: '${siteUrl}'}`
          };    
            
            props.context.httpClient.post(functionUrl, HttpClient.configurations.v1, postOptions).then((response) =>{     
             console.log(response) 
             console.log(response.nativeResponse.status)  
             //window.location.reload(true);
             if(response.nativeResponse.status == 200 ){
               getItems();
               sp.web.lists.getByTitle("syncGroupAppSettings").items.get().then(items => {
                items.forEach(item => {
                    sp.web.lists.getByTitle("syncGroupAppSettings").items.getById(item.ID).update({
                        SecurityGroupID: props.ID,
                        SecurityGroupName : props.securityGroupName
                    })
                });
            })
             }
            })    
            
                .catch ((response: any) => {    
                let errMsg: string = `WARNING - error when calling URL ${functionUrl}. Error = ${response.message}`;    
                console.log(errMsg)
              });       
      }    


    function AddGroup(){
      props.setProgress(true)
      callAzureFunction();
      // props.context.msGraphClientFactory
      // .getClient()
      // .then((client: MSGraphClient) => {
      //   // get information about the current user from the Microsoft Graph
      //   client
      //   .api("/groups/{"+props.ID+"}")
      //     .get((error, response: any, rawResponse?: any) => {
      //       // handle the response
      //      // console.log(JSON.stringify(response));
      //       var responseJson = JSON.stringify(response)
      //       var responseParsed =JSON.parse(responseJson)
      //      // console.log("adding group")
      //      // console.log(responseParsed)
      //       sp.web.lists.getByTitle("syncGroupAppSettings").items.get().then(items => {
      //           items.forEach(item => {
      //               sp.web.lists.getByTitle("syncGroupAppSettings").items.getById(item.ID).update({
      //                   SecurityGroupID: props.ID,
      //                   SecurityGroupName : responseParsed.displayName
      //               })
      //           });
      //       })
      //     })
      //   })
     
    }
    var disabled = props.ID == ""  || props.progress ? true : false;

    return (
      <PrimaryButton className={styles.addButton} onClick={AddGroup} disabled={disabled} text="Add" />
    )    
}