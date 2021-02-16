$fullTenantName = "o365secondtenantobvie.onmicrosoft.com"
$tenantName = "o365secondtenantobvie"
$tenantId = "9dadc14c-6b46-4a9b-bf3b-ddab7bc5161d"
$azureAppName = "SyncGroupTest6"
$functionAppName = "PowershellOperation6"
$storageAccount = "storageaccountsync"
$resourceGroup = "SecondTenantTestResourceGroup"
$subscriptionId = "69379d11-2a14-4dd1-8df6-32f1dae0e50f"
$siteUrl = "https://o365secondtenantobvie.sharepoint.com/sites/ThirdGroup"

$spadmin = "https://o365secondtenantobvie-admin.sharepoint.com"

$azureAppId = "6fe99093-c0d5-4554-b729-c2cd925fa59b"
$thumbprint = "8B23F2E877251D4CC90A38EE4508F1B7A34404D3"




.\deployAzure.ps1 -FullTenantName $fullTenantName -TenantName $tenantName -TenantId $tenantId -AzureAppName $azureAppName -FunctionAppName $functionAppName -StorageAccount $storageAccount -SubscriptionId $subscriptionId -ResourceGroupName $resourceGroup

.\deploySPFX.ps1 -SiteUrl $siteUrl -TenantName $tenantName -FunctionAppName $functionAppName -AzureAppId $azureAppId -TenantId $tenantId -Thumbprint $thumbprint

Connect-SPOService -Url $spadmin
Set-SPOsite $siteUrl -DenyAddAndCustomizePages 0

Set-SPOTenantCdnEnabled -CdnType Both -Enable $true

Connect-PnPOnline -Url $siteUrl -UseWebLogin


Install-PnPApp -Identity A2ED129B-2C98-4CDC-BC7D-7CE95A2EF468