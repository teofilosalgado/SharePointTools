Connect-PnPOnline -Url 'https://digitalfortytwo-admin.sharepoint.com/' -Interactive;

$DenyAddAndCustomizePagesStatusEnum = [Microsoft.Online.SharePoint.TenantAdministration.DenyAddAndCustomizePagesStatus]

$context = Get-PnPContext
$site = Get-PnPTenantSite -Detailed -Url 'https://digitalfortytwo.sharepoint.com/sites/PnPDestino'
$site.DenyAddAndCustomizePages = $DenyAddAndCustomizePagesStatusEnum::Disabled
$site.Update()
$context.ExecuteQuery()

# This row should output list of your sites' URLs and the status of their custom scripting (disabled or not)
Get-PnPTenantSite -Detailed -Url 'https://digitalfortytwo.sharepoint.com/sites/PnPDestino' | Select-Object url, DenyAddAndCustomizePages

Disconnect-PnPOnline
