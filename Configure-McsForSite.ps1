param (
    [Parameter(Mandatory=$true)]
    [string]$siteUrl,
    [Parameter(Mandatory=$true)]
    [string]$botUrl,
    [Parameter(Mandatory=$true)]
    [string]$botName,
    [Parameter(Mandatory=$true)]
    [string]$customScope,
    [Parameter(Mandatory=$true)]
    [string]$clientId,
    [Parameter(Mandatory=$true)]
    [string]$authority,
    [Parameter(Mandatory=$true)]
    [string]$buttonLabel,
    [Parameter(Mandatory=$true)]
    [switch]$greet
)
##### Replace Clientid with your PNP PowerShell App Registration Client ID
Connect-PnPOnline -Url $siteUrl -Interactive -ClientId PNP PowerShell App Registration Client ID
$action = (Get-PnPCustomAction | Where-Object { $_.Title -eq "PvaSso" })[0]
$action.ClientSideComponentProperties = @{
    "botURL" = $botUrl
    "customScope" = $customScope
    "clientID" = $clientId
    "authority" = $authority
    "greet" = $greet.isPresent
    "buttonLabel" = $buttonLabel
    "botName" = $botName
} | ConvertTo-Json -Compress
$action.Update()
Invoke-PnPQuery

