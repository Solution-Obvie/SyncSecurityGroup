Param(
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
            
            $app = Initialize-PnPPowerShellAuthentication -ApplicationName $AzureAppName -Tenant $TenantName -OutPath .\certificates -CertificatePassword (ConvertTo-SecureString -String "MyPassword" -AsPlainText -Force)
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
$pnpConnect = Connect-PnPOnline -Url $tenantAdminUrl -Credentials (Get-Credential)


#Registering the Azure Ad Application
#CreateAzureAdApp

#Deploy Azure Function*
#az functionapp create -g "SPFX_SyncGroup_RG" -n "functionAppSyncTest" -s "storageaccountspfxs8dd7" --consumption-plan-location "westus" --functions-version 3 --runtime "powershell" --runtime-version "7.0" --subscription "d6bb92ec-09b1-468a-b1b4-3460076686e4"
az functionapp config appsettings set --name "functionAppSyncTest" --resource-group "SPFX_SyncGroup_RG" --settings "AdminSharePointSite=$global:tenantAdminUrl"
az functionapp config appsettings set --name "functionAppSyncTest" --resource-group "SPFX_SyncGroup_RG" --settings "AzureAppId=c5d7f56a-bad4-4adb-b606-a8b0cd4aa9bf"
az functionapp config appsettings set --name "functionAppSyncTest" --resource-group "SPFX_SyncGroup_RG" --settings "TenantId=$TenantId"
az functionapp config appsettings set --name "functionAppSyncTest" --resource-group "SPFX_SyncGroup_RG" --settings "WEBSITE_USE_ZIP=1"
az functionapp config appsettings set --name "functionAppSyncTest" --resource-group "SPFX_SyncGroup_RG" --settings "WEBSITE_LOAD_CERTIFICATES=730D95220B36B414BF0422EE0FF9A7BC634A121A"

az functionapp config ssl upload --certificate-file './certificates/SyncGroupTest.pfx' --certificate-password "MyPassword" --name "functionAppSyncTest" --resource-group "SPFX_SyncGroup_RG"