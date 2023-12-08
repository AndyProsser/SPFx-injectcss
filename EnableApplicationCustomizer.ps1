# Change these variables to enable the extension
$tenantUrl = "https://teachforaustralia.sharepoint.com/sites/community"
$customCSSUrl = "$tenantUrl/Style%20Library/custom.css"

# Shorten name!
$DenyAddAndCustomizePagesStatusEnum = [Microsoft.Online.SharePoint.TenantAdministration.DenyAddAndCustomizePagesStatus]

# Get credentials
#$credentials = Get-Credential
Connect-PnPOnline $tenantUrl -Interactive

# Connect to tenant
$context = Get-PnPContext
$web = Get-PnPWeb
$context.Load($web)
Invoke-PnPQuery

# Deploy custom action
$ca = $web.UserCustomActions.Add()
$ca.ClientSideComponentId = "5a1fcffd-dfeb-4844-b478-1feb4325a5a7"
$ca.ClientSideComponentProperties = "{""cssurl"":""$customCSSUrl""}"
$ca.Location = "ClientSideExtension.ApplicationCustomizer"
$ca.Name = "InjectCssApplicationCustomizer"
$ca.Title = "Inject CSS Application Extension"
$ca.Description = "Injects custom CSS to make minor style modifications to SharePoint Online"
$ca.Update()

$context.Load($web.UserCustomActions)
Invoke-PnPQuery

# Allow Access to Style Library
$site = Get-PnPTenantSite -Detailed -Url $tenantUrl
$site.DenyAddAndCustomizePages = $DenyAddAndCustomizePagesStatusEnum::Disabled
$site.Update()
$context.ExecuteQuery()

$status = $null
Write-Host "Waiting..." -NoNewline
do {
  Write-Host "." -NoNewline
  Start-Sleep -Seconds 5
  $site = Get-PnPTenantSite -Detailed -Url $tenantUrl
  $status = $site.Status

} While ($status -ne 'Active')
Write-Host $status -ForegroundColor Green
