using namespace System.Collections.Generic;

function Export-SharePointTemplate {
    <#
    .SYNOPSIS
        Export SharePoint site template.
    .DESCRIPTION
        This script export every single page in a SharePoint site to the 
        $OutputPath folder which, by the end of the execution, will contain all
        the templates organized by Guid and a Log.csv index file.
    .EXAMPLE
        PS C:\> Export-SharePointTemplate -SourceUrl "mytenant.sharepoint.com" -OutputPath ".\Result"
    .PARAMETER SourceUrl
        Specify the Url of the source SharePoint site.
    .PARAMETER OutputPath
        Specify the location to store the templates and index file.
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
        Write-Verbose "[Export-SharePointTemplate] Start!";
        Write-Verbose "[Export-SharePointTemplate] Connecting to $($SourceUrl)";
        Connect-PnPOnline -Url $SourceUrl -UseWebLogin -WarningAction Ignore;
        
        Write-Verbose "[Export-SharePointTemplate] Creating $($OutputPath) folder";
        if (-not (Test-Path $OutputPath)) {
            New-Item $OutputPath -ItemType Directory | Out-Null;   
        }
        $ResultPath = Resolve-Path $OutputPath;

        Write-Verbose "[Export-SharePointTemplate] Querying $($SourceUrl) pages";
        $PageList = Get-PnPListItem -List "SitePages";
        $Records = [List[PSCustomObject]]::new();
    }
    
    process {
        $CurrentTemplateIndex = 0;
        foreach ($PageItem in $PageList) {
            $PageName = $PageItem.FieldValues["FileLeafRef"];
            $PageGuid = $PageItem.FieldValues["GUID"];
            $TemplateFolderPath = Join-Path -Path $ResultPath -ChildPath $PageGuid;
            $TemplateFilePath = Join-Path -Path $TemplateFolderPath -ChildPath "Template.xml";
            
            if (-not (Test-Path $TemplateFolderPath)) {
                New-Item -ItemType directory -Path $TemplateFolderPath | Out-Null;
            }

            Write-Progress `
                -Activity "Exporting $($PageName)" `
                -Status "$($CurrentTemplateIndex) of $($PageList.Count) pages" `
                -PercentComplete (($CurrentTemplateIndex / $PageList.Count) * 100);

            try {
                Export-PnPClientSidePage `
                    -Identity $PageName `
                    -Out $TemplateFilePath `
                    -PersistBrandingFiles `
                    -ErrorAction Stop `
                    -Force;
                
                $Time = Get-Date -Format "o";
                $Record = [PSCustomObject]@{
                    Guid = $PageGuid
                    Name = $PageName
                    Path = $TemplateFilePath
                    Time = $Time
                };
                $Records.Add($Record);
                $CurrentTemplateIndex = $CurrentTemplateIndex + 1;
            }
            catch {
                Write-Error "[Export-SharePointTemplate] Error exporting page $($PageName)!";
            }
        }
        Write-Progress `
            -Activity "Exporting customization" `
            -PercentComplete 0;

        $CustomizationTemplateFolderPath = Join-Path -Path $ResultPath -ChildPath "Customization";
        $CustomizationTemplateFilePath = Join-Path -Path $CustomizationTemplateFolderPath -ChildPath "Customization.xml";

        Get-PnPProvisioningTemplate `
            -Handlers Navigation `
            -Out $CustomizationTemplateFilePath `
            -Force;
            
        $Time = Get-Date -Format "o";
        $Record = [PSCustomObject]@{
            Guid = "Customization"
            Name = "Customization"
            Path = $CustomizationTemplateFilePath
            Time = $Time
        };
        $Records.Add($Record);

        Write-Progress `
            -Activity "Exporting customization" `
            -PercentComplete 100;
    }
    
    end {
        $LogPath = Join-Path -Path $ResultPath -ChildPath "Log.csv";
        Write-Verbose "[Export-SharePointTemplate] Saving report to $($LogPath)"
        $Records | Export-Csv -Path $LogPath -Force;
        Write-Verbose "[Export-SharePointTemplate] Done!";
    }
}

Export-ModuleMember -Function Export-SharePointTemplate;