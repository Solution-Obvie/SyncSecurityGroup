import { HttpClient, SPHttpClient, HttpClientConfiguration, HttpClientResponse, ODataVersion, IHttpClientConfiguration, IHttpClientOptions, ISPHttpClientOptions } from '@microsoft/sp-http'; 

export function callAzureFunction(functionUri, context, microsoftGroupID, securitygroupID) {
    return new Promise((resolve) => {

        const requestHeaders: Headers = new Headers();    
        requestHeaders.append("Content-type", "application/json");  
        requestHeaders.append("Cache-Control", "no-cache");    
          
        var siteUrl: string = context.pageContext.web.absoluteUrl; 
        
        const postOptions: IHttpClientOptions = {    
            headers: requestHeaders,
            body:`{ microsoftgroupID:  '${microsoftGroupID}', securitygroupID:  '${securitygroupID}', siteUrl: '${siteUrl}'}`
        };   
     
          
          context.httpClient.post(functionUri , HttpClient.configurations.v1, postOptions).then((response) =>{     
           console.log(response) 
           console.log(response.nativeResponse.status) 
           resolve({"status": response.nativeResponse.status})
          })    
          
              .catch ((response: any) => {    
              let errMsg: string = `WARNING - error when calling URL ${functionUri}. Error = ${response.message}`;    
              console.log(errMsg)
            });   

    })
}