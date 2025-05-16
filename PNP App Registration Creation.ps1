# 1. Install & Import AzureAD module (if needed)
# ─────────────────────────────────────────────────────────────────────────────
if (-not (Get-Module -ListAvailable -Name AzureAD)) {
    Install-Module -Name AzureAD -Scope CurrentUser -Force
}
Import-Module AzureAD

# ─────────────────────────────────────────────────────────────────────────────
# 2. Connect to Azure AD
# ─────────────────────────────────────────────────────────────────────────────
Connect-AzureAD

# ─────────────────────────────────────────────────────────────────────────────
# 3. Define variables
# ─────────────────────────────────────────────────────────────────────────────
$displayName   = 'PnP PowerShell App'
# A unique URI for your app; you can change to match your organization’s domain
$identifierUri = "https://contoso.com/apps/pnppowershell"
$secretDesc    = 'PnPSecret'
$startDate     = (Get-Date).ToUniversalTime()
$endDate       = $startDate.AddYears(1)

# ─────────────────────────────────────────────────────────────────────────────
# 4. Create the App Registration
# ─────────────────────────────────────────────────────────────────────────────
$app = New-AzureADApplication `
    -DisplayName     $displayName `
    -IdentifierUris  @($identifierUri) `
    -PasswordCredentials @() `
    -ReplyUrls       @()  # no web redirect required for client-credentials

Write-Host "➜ Created AAD Application:"
Write-Host "    Name: $($app.DisplayName)"
Write-Host "    AppId: $($app.AppId)"
Write-Host "    ObjectId: $($app.ObjectId)"
Write-Host ''

# ─────────────────────────────────────────────────────────────────────────────
# 5. Create its Service Principal
# ─────────────────────────────────────────────────────────────────────────────
$sp = New-AzureADServicePrincipal -AppId $app.AppId

Write-Host "➜ Created Service Principal:"
Write-Host "    SP ObjectId: $($sp.ObjectId)"
Write-Host ''

# ─────────────────────────────────────────────────────────────────────────────
# 6. Add a client secret
# ─────────────────────────────────────────────────────────────────────────────
$secret = New-AzureADApplicationPasswordCredential `
    -ObjectId              $app.ObjectId `
    -CustomKeyIdentifier   $secretDesc `
    -StartDate             $startDate `
    -EndDate               $endDate

# The returned object’s `Value` property is the secret’s plaintext
$secretValue = $secret.Value
Write-Host "➜ Created client secret (copy this now, it won’t be retrievable again):"
Write-Host "    Description: $secretDesc"
Write-Host "    Expires:     $($secret.EndDate)"
Write-Host "    SecretValue: $secretValue"
Write-Host ''

# ─────────────────────────────────────────────────────────────────────────────
# 7. Grant SharePoint App-Only Permission: Sites.FullControl.All
# ─────────────────────────────────────────────────────────────────────────────
# 7.1 Retrieve the SharePoint service principal
$sharePointSp = Get-AzureADServicePrincipal `
    -Filter "AppId eq '00000003-0000-0ff1-ce00-000000000000'"

# 7.2 Find the AppRole for Sites.FullControl.All
$appRole = $sharePointSp.AppRoles `
    | Where-Object {
        ($_ .Value -eq 'Sites.FullControl.All') -and
        ($_ .AllowedMemberTypes -contains 'Application')
      }

if (-not $appRole) {
    Throw "Unable to find the Sites.FullControl.All AppRole on SharePoint service principal."
}

# 7.3 Assign that role to your service principal
New-AzureADServiceAppRoleAssignment `
    -ObjectId    $sp.ObjectId `
    -PrincipalId $sp.ObjectId `
    -ResourceId  $sharePointSp.ObjectId `
    -Id          $appRole.Id

Write-Host "➜ Assigned ‘Sites.FullControl.All’ to service principal."
Write-Host ''

# ─────────────────────────────────────────────────────────────────────────────
# 8. Summary: connection parameters for PnP.PowerShell
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "✔ Setup complete. Use these values to connect in app-only mode:`n"
Write-Host "   Tenant ID:   $(Get-AzureADTenantDetail).ObjectId"
Write-Host "   Client ID:   $($app.AppId)"
Write-Host "   Client Secret:$secretValue"

