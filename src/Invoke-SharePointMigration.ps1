[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string]
    $SourceUrl,

    [Parameter(Mandatory)]
    [string]
    $DestinationUrl,

    [Parameter(Mandatory = $false)]
    [switch]
    $ConvertToTeamSite = $false
)

Import-Module "$PSScriptRoot\Export-SharePointTemplate.psm1";
Import-Module "$PSScriptRoot\Import-SharePointTemplate.psm1";

$OutputPath = Resolve-Path ".\Result";
if (-not (Test-Path $OutputPath)) {
    New-Item $OutputPath -ItemType Directory | Out-Null;   
}
$TemplatePath = Resolve-Path $OutputPath;

Write-Verbose "[Invoke-Migration] Start!";
Export-SharePointTemplate -SourceUrl $SourceUrl -OutputPath $TemplatePath;
Import-SharePointTemplate -DestinationUrl $DestinationUrl -InputPath $TemplatePath -ConvertToTeamSite $ConvertToTeamSite;
Remove-Item -Recurse -Force $TemplatePath;
Write-Verbose "[Invoke-Migration] Done!";