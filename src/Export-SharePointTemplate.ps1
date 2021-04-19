using namespace System.Collections.Generic;

function Export-SharePointTemplate {
    <#
    .SYNOPSIS
        Export SharePoint site template.
    .DESCRIPTION
        This script export every single page in a SharePoint site to a Result folder 
        containing the templates organized by Guid.
    .EXAMPLE
        Export-SharePointTemplate -SourceUrl "mytenant.sharepoint.com" -OutputPath ".\Result"
    .PARAMETER SourceUrl
        Specify the Url of the SharePoint site.
    .PARAMETER OutputPath
        Specify the location to store the templates.
    .LINK
        https://github.com/teofilosalgado/SharePointTools
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $SourceUrl,

        [Parameter(Mandatory)]
        [string]
        $OutputPath
    )
    
    begin {
        Write-Verbose "[Export-SharePointTemplate] Start!"
        Write-Verbose "[Export-SharePointTemplate] Connecting to $($SourceUrl)"
        Connect-PnPOnline -Url $SourceUrl -UseWebLogin -WarningAction Ignore;
        
        Write-Verbose "[Export-SharePointTemplate] Creating $($OutputPath) folder"
        if (-not (Test-Path $OutputPath)) {
            New-Item $OutputPath -ItemType Directory | Out-Null;   
        }
        $ResultPath = Resolve-Path $OutputPath;
    }
    
    process {
        $PageList = Get-PnPListItem -List "SitePages";
        $Records = [List[PSCustomObject]]::new();

        $Index = 0;
        foreach ($PageItem in $PageList) {
            $Name = $PageItem.FieldValues["FileLeafRef"];
            $Guid = $PageItem.FieldValues["GUID"];
            $TemplateFolderPath = Join-Path -Path $ResultPath -ChildPath $Guid;
            $TemplateFilePath = Join-Path -Path $TemplateFolderPath -ChildPath "Template.xml";
            
            if (-not (Test-Path $TemplateFolderPath)) {
                New-Item -ItemType directory -Path $TemplateFolderPath | Out-Null;
            }

            Write-Progress `
                -Activity "Exporting $($Name).." `
                -Status "$($Index) of $($PageList.Count) pages" `
                -PercentComplete (($Index / $PageList.Count) * 100);

            try {
                Export-PnPClientSidePage `
                    -Identity $Name `
                    -Out $TemplateFilePath `
                    -PersistBrandingFiles `
                    -ErrorAction Stop `
                    -Force;
                
                $Time = Get-Date -Format "o";
                $Record = [PSCustomObject]@{
                    Guid = $Guid
                    Name = $Name
                    Path = $TemplateFilePath
                    Time = $Time
                }
                $Records.Add($Record);
                $Index = $Index + 1;
            }
            catch {
                Write-Error "Error exporting page: $($Name)!";
            }
        }
    }
    
    end {
        $LogPath = Join-Path -Path $ResultPath -ChildPath "Log.csv";
        Write-Verbose "[Export-SharePointTemplate] Saving report to $($LogPath)"
        $Records | Export-Csv -Path $LogPath -Force;
        Write-Verbose "[Export-SharePointTemplate] Done!"
    }
}