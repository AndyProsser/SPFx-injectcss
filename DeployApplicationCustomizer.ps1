# Use this file to deploy the extension to your application catalog
$tenantUrl = "https://<your-tenant>.sharepoint.com"

# Get credentials
$credentials = Get-Credential
Connect-PnPOnline $tenantUrl -Credentials $credentials

Add-PnPApp -path .\sharepoint\solution\spfx-injectcss.sppkg -Publish -SkipFeatureDeployment -Overwrite
