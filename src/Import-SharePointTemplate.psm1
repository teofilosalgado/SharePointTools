function Import-SharePointTemplate {
    <#
    .SYNOPSIS
        Import SharePoint site template.
    .DESCRIPTION
        This script import every template stored in the $OutputPath folder
        to the desired SharePoint site at $DestinationUrl. Optionally, if the
        destination site is a Team Site and the origin site was an Organization
        Site you may use the ConvertToTeamSite flag to prevent errror.s
    .EXAMPLE
        PS C:\> Import-SharePointTemplate -DestinationUrl "mytenant.sharepoint.com" -InputPath ".\Result" -ConvertToTeamSite
    .PARAMETER DestinationUrl
        Specify the Url of the destination SharePoint site.
    .PARAMETER OutputPath
        Specify the location where the templates are stored.
    .PARAMETER ConvertToTeamSite
        Convert the template to fit a Team Site if the origin site was an
        Organization site.
    .LINK
        https://github.com/teofilosalgado/SharePointTools
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $DestinationUrl,

        [Parameter(Mandatory)]
        [string]
        $InputPath,

        [Parameter(Mandatory = $false)]
        [switch]
        $ConvertToTeamSite = $false
    )
    
    begin {
        Write-Verbose "[Import-SharePointTemplate] Start!";
        Write-Verbose "[Import-SharePointTemplate] Connecting to $($DestinationUrl)";
        Connect-PnPOnline -Url $DestinationUrl -UseWebLogin -WarningAction Ignore;
        
        $CsvLocation = Join-Path -Path $InputPath -ChildPath "Log.csv";
        Write-Verbose "[Import-SharePointTemplate] Reading report at $($CsvLocation)";
        $Templates = Import-Csv -Path $CsvLocation;
        $Total = $Templates.Count;
    }
    
    process {
        $CurrentTemplateIndex = 0;
        foreach ($Template in $Templates) {
            $PageName = $Template.Name;
            $TemplatePath = $Template.Path;
    
            if ($ConvertToTeamSite) {
                $FolderPath = Split-Path -Path $TemplatePath
                $AltTemplatePath = Join-Path -Path $FolderPath -ChildPath "AltTemplate.xml";
                Copy-Item $TemplatePath -Destination $AltTemplatePath;
                (Get-Content $AltTemplatePath).replace('OneColumnFullWidth', 'OneColumn') | Set-Content $AltTemplatePath;
                $TemplatePath = $AltTemplatePath;
            }
            try {
                Invoke-ApplySharePointTemplate `
                    -DestinationUrl $DestinationUrl `
                    -PageName $PageName `
                    -TemplatePath $TemplatePath `
                    -Total $Total `
                    -Index $CurrentTemplateIndex `
                    -ErrorAction Stop;
            }
            catch {
                Write-Error "[Import-SharePointTemplate] Error importing page $($PageName)!";
            }
            $CurrentTemplateIndex = $CurrentTemplateIndex + 1;
        }
    }
    
    end {
        Write-Verbose "[Import-SharePointTemplate] Done!";
    }
}

function Invoke-ApplySharePointTemplate {
    param (
        [Parameter(Mandatory)]
        [string]
        $DestinationUrl,
        
        [Parameter(Mandatory)]
        [string]
        $PageName,

        [Parameter(Mandatory)]
        [string]
        $TemplatePath,

        [Parameter(Mandatory)]
        [int]
        $Total,

        [Parameter(Mandatory)]
        [int]
        $Index
    )

    $Count = 0;
    do {
        $Count = $Count + 1;
        try {
            Write-Progress `
                -Activity "Importing $($PageName). Attempt $($Count) of 3" `
                -Status "$($Index) of $($Total) pages" `
                -PercentComplete (($Index / $Total) * 100);
            Apply-PnPProvisioningTemplate -Path $TemplatePath -ClearNavigation;
            return;
        }
        catch {
            Connect-PnPOnline -Url $DestinationUrl -UseWebLogin;
            Start-Sleep -Seconds 30;
        }
    } while ($Count -lt 3)
    throw Get-PnPException;
}

Export-ModuleMember -Function Import-SharePointTemplate;