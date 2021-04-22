using namespace System;
using namespace System.Web;

function Repair-WebParts {
    <#
    .SYNOPSIS
        Repair broken file viewer webparts.
    .DESCRIPTION
        This scripts looks for broken file viewer webparts in every single page
        of a given $SiteUrl and repairs it using embed iframes.
    .EXAMPLE
        PS C:\> Repair-WebParts -SiteUrl "mytenant.sharepoint.com"
    .PARAMETER SiteUrl
        Specify the Url of the SharePoint site.
    .LINK
        https://github.com/teofilosalgado/SharePointTools
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $SiteUrl
    )
    
    begin {
        Write-Verbose "[Repair-WebParts] Start!";
        Write-Verbose "[Repair-WebParts] Connecting to $($DestinationUrl)";
        Connect-PnPOnline -Url $SiteUrl -UseWebLogin -WarningAction Ignore;

        $BaseUrl = "https://" + ([Uri]$SiteUrl).Host;
        $PageList = Get-PnPListItem -List "SitePages";
    }
    
    process {
        foreach ($PageItem in $PageList) {
            $PageName = $PageItem.FieldValues["FileLeafRef"];
            $Page = Get-PnPClientSidePage -Identity $PageName;
            $Controls = $Page.Controls.Where( { $_.Type.Name -eq "PageWebPart" -and $_.Title -eq "File viewer" });

            $CurrentControlIndex = 0;
            foreach ($Control in $Controls) {
                Write-Progress `
                    -Activity "Fixing $($PageName)" `
                    -Status "$($CurrentControlIndex) of $($Controls.Count) webparts" `
                    -PercentComplete (($CurrentControlIndex / $Controls.Count) * 100);

                $Found = $Control.HtmlPropertiesData -match 'href=["]([^"]+?)["]';
                if ($Found) {
                    $Url = $BaseUrl + $Matches[1];
                    $DecodedUrl = [HttpUtility]::UrlDecode($Url);
                    $Path = $([Uri]$DecodedUrl).LocalPath
                    $File = Get-PnPFile -Url $Path -AsListItem;
                    $Guid = $File.FieldValues["UniqueId"].Guid.ToString();
                    $Code = '<iframe width="400" height="250" frameborder="0" scrolling="no" src="' + $SiteUrl + '/_layouts/15/Doc.aspx?sourcedoc={' + $Guid + '}&action=embedview&wdAllowInteractivity=False&wdHideGridlines=True&wdHideHeaders=True&wdDownloadButton=True&wdInConfigurator=True"></iframe>';
        
                    $Column = $Control.Column.Order;
                    $Section = $Control.Section.Order;
                    $Order = $Control.Order;
        
                    $Control.Delete();
                            
                    Add-PnPPageWebPart `
                        -Column $Column `
                        -Section $Section `
                        -Order $Order `
                        -Page $Page `
                        -DefaultWebPartType ContentEmbed `
                        -WebPartProperties @{"embedCode" = $Code };
                    $Page.Save();
                }

                $CurrentControlIndex = $CurrentControlIndex + 1;
            }
        }
    }
    
    end {
        Write-Verbose "[Repair-WebParts] Done!";
    }
}

