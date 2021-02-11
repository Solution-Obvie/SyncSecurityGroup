Param(
    [Parameter(Mandatory = $true)]
    [String]
    $SiteUrl,
    [Parameter(Mandatory = $true)]
    [String]
    $TenantName,
    [Parameter(Mandatory = $true)]
    [String]
    $FunctionAppName

)


Connect-PnPOnline -Url $SiteUrl -Credentials (Get-Credential)
#deploy propertybag
Set-PnPPropertyBagValue -Key "MicrosoftGroupUsers" -Value " "
Set-PnPPropertyBagValue -Key "SecurityGroupUsers" -Value " "
Set-PnPPropertyBagValue -Key "MicrosoftGroup" -Value " "
Set-PnPPropertyBagValue -Key "LastSync" -Value " "
Set-PnPPropertyBagValue -Key "syncGroupAppEnabled" -Value "true"
Set-PnPPropertyBagValue -Key "AddedMember" -Value " "
Set-PnPPropertyBagValue -Key "RemovedMember" -Value " "
Set-PnPPropertyBagValue -Key "SecurityGroupLinked" -Value " "
Set-PnPPropertyBagValue -Key "FunctionAppAzure" -Value $FunctionAppName


$securityGroupsExcel = Import-Excel ".\SecurityGroups.xlsx"
$securityGroups = @()

foreach ($group in $securityGroupsExcel) {
    $securityGroups += @{Id = $group.Id ; Name = $group.Name } 
}

$securityGroupsJson = $securityGroups | ConvertToJson

Set-PnPPropertyBagValue -Key "SecurityGroups" -Value $securityGroupsJson 

#deploy app to web site
Connect-PnPOnline -Url "https://"$TenantName".sharepoint.com" -Credentials (Get-Credential)
Add-PnPApp -Path sharepoint\solution\sync-group-app.sppkg -Scope Site -Publish