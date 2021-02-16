Param(
    [Parameter(Mandatory = $true)]
    [String]
    $FullTenantName,
    [Parameter(Mandatory = $true)]
    [String]
    $TenantName,
    [Parameter(Mandatory = $true)]
    [String]
    $TenantId,
    [Parameter(Mandatory = $true)]
    [String]
    $AzureAppName,
    [Parameter(Mandatory = $true)]
    [String]
    $FunctionAppName,
     [Parameter(Mandatory = $true)]
    [String]
    $StorageAccount,
     [Parameter(Mandatory = $true)]
    [String]
    $SubscriptionId,
    [Parameter(Mandatory = $true)]
    [String]
    $CertificatePassword,
    [Parameter(Mandatory = $true)]
    [String]
    $ResourceGroupName
)

Add-Type -AssemblyName System.Web

if (Get-InstalledModule  -Name "microsoft.online.sharepoint.powershell") {
Write-Host "SharePoint Online Powershell installed"
}
else{
Install-Module microsoft.online.sharepoint.powershell -Scope CurrentUser
}

if (Get-InstalledModule  -Name "SharePointPnPPowerShellOnline") {
Write-Host "SharePoint PnP Powershell Online installed"
}
else{
Install-Module SharePointPnPPowerShellOnline -Scope CurrentUser -MinimumVersion 3.19.2003.0 -Force
}

if (Get-InstalledModule  -Name "ImportExcel") {
Write-Host "Import Excel installed"
}
else{
Install-Module ImportExcel -Scope CurrentUser
}

if (Get-InstalledModule  -Name "Az") {
Write-Host "Az installed"
}
else{
Install-Module Az -AllowClobber -Scope CurrentUser
}

# Global variables
$global:azureAppId = $null
$global:certificatThumbprint = $null
$global:tenantUrl = "https://$tenantName.sharepoint.com"
$global:tenantAdminUrl = "https://$tenantName-admin.sharepoint.com"

function GetAzureADApp {
    param ($appName)

    $app = az ad app list --filter "displayName eq '$appName'" | ConvertFrom-Json

    return $app

}
function CreateAzureAdApp {
  
    try {
        Write-Host "### AZURE AD APP CREATION ###" -ForegroundColor Yellow

        # Check if the app already exists - script has been previously executed
        $app = GetAzureADApp $AzureAppName

        if (-not ([string]::IsNullOrEmpty($app))) {

           
            # Update azure ad app registration using CLI
            Write-Host "Azure AD App '$AzureAppName' already exists - updating existing app..." -ForegroundColor Yellow
            $global:azureAppId = $app.appId
            az ad app update --id $global:azureAppId --required-resource-accesses './manifest.json' 

            Write-Host "Waiting for app to finish updating..."

            Start-Sleep -s 60

            Write-Host "Updated Azure AD App" -ForegroundColor Green

        } 
        else {
            # Create the app
            Write-Host "No Azure Ad app found" -ForegroundColor Yellow
            
            $app = Initialize-PnPPowerShellAuthentication -ApplicationName $AzureAppName -Tenant $FullTenantName -OutPath .\certificates -CertificatePassword (ConvertTo-SecureString -String $CertificatePassword -AsPlainText -Force)
            #
            $global:certificatThumbprint = $app.'Certificate Thumbprint'
            $global:azureAppId = $app.AzureAppId
            az ad app update --id $global:azureAppId --required-resource-accesses './manifest.json' 



        }
        #$global:certificatThumbprint = $app.'Certificate Thumbprint'

        Write-Host "Granting admin consent for Microsoft Graph..." -ForegroundColor Yellow

        # Grant admin consent for app registration required permissions using CLI
        az ad app permission admin-consent --id $global:azureAppId
        
        Write-Host "Waiting for admin consent to finish..."

        Start-Sleep -s 60
        
        Write-Host "Granted admin consent" -ForegroundColor Green

        # Get service principal id for the app we created

        Write-Host "### AZURE AD APP CREATION FINISHED ###" -ForegroundColor Green
    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Host "Error occured while creating an Azure AD App: $errorMessage" -ForegroundColor Red
    }
}




#Connections 
Connect-AzureAD
$cliLogin = az login --allow-no-subscriptions
$pnpConnect = Connect-PnPOnline -Url $tenantUrl -Credentials (Get-Credential)


#Registering the Azure Ad Application
CreateAzureAdApp

#Deploy Azure Function*
#Create function app
az functionapp create -g  $ResourceGroupName -n $FunctionAppName -s $StorageAccount --consumption-plan-location "westus" --functions-version 3 --runtime "powershell" --runtime-version "7.0" --subscription $SubscriptionId
#Adding app settings
az functionapp config appsettings set --name $FunctionAppName --resource-group $ResourceGroupName --settings "AdminSharePointSite=$global:tenantAdminUrl"
az functionapp config appsettings set --name $FunctionAppName --resource-group $ResourceGroupName --settings "AzureAppId= $global:azureAppId"
az functionapp config appsettings set --name $FunctionAppName --resource-group $ResourceGroupName --settings "TenantId=$TenantId"
#az functionapp config appsettings set --name $FunctionAppName --resource-group $ResourceGroupName --settings "WEBSITE_USE_ZIP=1"
az functionapp config appsettings set --name $FunctionAppName --resource-group $ResourceGroupName --settings "WEBSITE_LOAD_CERTIFICATES= $global:certificatThumbprint"
#Uploading certificate
$certificate = "./certificates/$AzureAppName.pfx"
az functionapp config ssl upload --certificate-file $certificate --certificate-password (ConvertTo-SecureString -String $CertificatePassword -AsPlainText -Force) --name $FunctionAppName --resource-group $ResourceGroupName
az functionapp cors add -g $ResourceGroupName -n $FunctionAppName --allowed-origins $global:tenantUrl
#Uploading scripts
az functionapp deployment source config-zip -g $ResourceGroupName -n $FunctionAppName --src .\PowerShellGroupOperation.zip

Add-PnPApp -Path "..\sharepoint\solution\sync-group-app.sppkg" -Publish