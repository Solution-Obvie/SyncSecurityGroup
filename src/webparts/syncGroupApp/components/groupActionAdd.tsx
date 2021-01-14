import * as React from 'react';
import styles from './SyncGroupApp.module.scss';
import { sp } from "@pnp/sp";
import { Item } from '@pnp/sp/items';
import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { HttpClient, SPHttpClient, HttpClientConfiguration, HttpClientResponse, ODataVersion, IHttpClientConfiguration, IHttpClientOptions, ISPHttpClientOptions } from '@microsoft/sp-http'; 
import { PrimaryButton } from 'office-ui-fabric-react';
import { getItem } from './functions/updateItem'
import {callAzureFunction} from './functions/callAzureFunction'

export default function GroupActionAdd(props) {


    var functionUri = "https://powershellgroupoperation.azurewebsites.net/api/AddSecurityGroup";    

    function AddGroup(){
      props.setProgress(true)
      callAzureFunction(functionUri, props.context, props.group.ID, props.ID)
      .then(data =>
        getItem()
        .then(data => {
          props.setGroup(data)
          props.setProgress(false)
        }
        )
      )  
    }
    var disabled = props.ID == ""  || props.progress ? true : false;

    return (
      <PrimaryButton className={styles.addButton} onClick={AddGroup} disabled={disabled} text="Add" />
    )    
}