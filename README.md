# SharePointTools
This script collection was intended to make SharePoint Online (SPO) provisioning easier, allowing for fast and reliable migrations.

## How it works?
Everything is based upon SharePointPnPPowerShellOnline commandlets, requiring way lower privileges over the instance than using PnP.PowerShell.

## How to run it?
After cloning/downloading this repository, you just need to import the module you want to use beforehand using `Import-Module` and you are good to go! Make sure you have installed a compatible version of SharePointPnPPowerShellOnline beforehand, you may find it [here](https://www.powershellgallery.com/packages/SharePointPnPPowerShellOnline).
The migration process normally takes 2 steps, exporting and importing the site template. Whenever a WebPart gets corrupted in the process you might want to use the repair module to fix the issue.
- ### Exporting
  This module export every single page in a SharePoint site to the $OutputPath folder which, by the end of the execution, will contain all the templates organized by Guid and a Log.csv index file.
    
    **Example:**
    
    ```PS C:\> Export-SharePointTemplate -SourceUrl "mytenant.sharepoint.com" -OutputPath ".\Result"```
        
    **Parameters:** 
      
    - `SourceUrl` Specify the Url of the source SharePoint site.
    - `OutputPath` Specify the location to store the templates and index (Log.csv) file.

- ### Importing
  This script import every template stored in the $OutputPath folder to the desired SharePoint site at $DestinationUrl. Optionally, if the destination site is a Team Site and the origin site was an Organization Site you may use the ConvertToTeamSite flag to prevent errors.
    
    **Example:**
    
    ```PS C:\> Import-SharePointTemplate -DestinationUrl "mytenant.sharepoint.com" -InputPath ".\Result" -ConvertToTeamSite```
        
    **Parameters:** 
      
    - `DestinationUrl` Specify the Url of the destination SharePoint site.
    - `OutputPath` Specify the location where the templates are stored.
    - `[ConvertToTeamSite]` Convert the template to fit a Team Site if the origin site was an Organization site.

- ### (Optional) Repairing WebParts
  This module repair broken file viewer webparts.
    
    **Example:**
    
    ```PS C:\> Repair-WebParts -SiteUrl "mytenant.sharepoint.com"```
        
    **Parameters:** 
      
    - `SiteUrl` Specify the Url of the SharePoint site.

## Minimum requirements:
 - SharePointPnPPowerShellOnline: 3.29.x
 - PowerShell: 5
