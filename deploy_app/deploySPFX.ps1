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
$m365Name = Get-PnPPropertyBag -key "GroupAlias"
$m365Id = Get-PnPPropertyBag -key "GroupId"
$MicrosoftGroup = @{Id = $m365Id ; Name = $m365Name}
$MicrosoftGroupJson = $MicrosoftGroup | ConvertTo-Json
$SecurityGroup = @{ Id = "" ; Name = "" }
$SecurityGroupJson = $SecurityGroup | ConvertTo-Json
#deploy propertybag
Set-PnPPropertyBagValue -Key "MicrosoftGroupUsers" -Value " "
Set-PnPPropertyBagValue -Key "SecurityGroupUsers" -Value " "
Set-PnPPropertyBagValue -Key "MicrosoftGroup" -Value $MicrosoftGroupJson
Set-PnPPropertyBagValue -Key "LastSync" -Value " "
Set-PnPPropertyBagValue -Key "syncGroupAppEnabled" -Value "true"
Set-PnPPropertyBagValue -Key "AddedMember" -Value " "
Set-PnPPropertyBagValue -Key "RemovedMember" -Value " "
Set-PnPPropertyBagValue -Key "SecurityGroupLinked" -Value $SecurityGroupJson
Set-PnPPropertyBagValue -Key "FunctionAppAzure" -Value $FunctionAppName


$securityGroupsExcel = Import-Excel ".\SecurityGroups.xlsx"
$securityGroups = @()

foreach ($group in $securityGroupsExcel) {
    $securityGroups += @{Id = $group.Id ; Name = $group.Name } 
}

$securityGroupsJson = $securityGroups | ConvertTo-Json 

Set-PnPPropertyBagValue -Key "SecurityGroups" -Value $securityGroupsJson 

#deploy app to web site
$url = "https://"+$TenantName+".sharepoint.com"
Connect-PnPOnline -Url $url -Credentials (Get-Credential)
Add-PnPApp -Path ..\sharepoint\solution\sync-group-app.sppkg -Scope Site -Publish