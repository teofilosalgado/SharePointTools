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

Import-Module "$PSScriptRoot\Export-SharePointTemplate.psm1" -Force;
Import-Module "$PSScriptRoot\Import-SharePointTemplate.psm1" -Force;

$TemplatePath = "..\Temp";

$ShouldConvertTeamSite = $false;
if ($ConvertToTeamSite) {
    $ShouldConvertTeamSite = $true;
}

Write-Verbose "[Invoke-SharePointMigration] Start!";
Export-SharePointTemplate -SourceUrl $SourceUrl -OutputPath $TemplatePath;
Import-SharePointTemplate -DestinationUrl $DestinationUrl -InputPath $TemplatePath -ConvertToTeamSite $ShouldConvertTeamSite;
Remove-Item -Recurse -Force $TemplatePath;
Write-Verbose "[Invoke-SharePointMigration] Done!";