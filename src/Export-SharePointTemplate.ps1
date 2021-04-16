using module ".\Item.psm1";

[CmdletBinding()]
param (
    # URL of the source site.
    [Parameter(Mandatory)]
    [string]
    $SourceUrl
)

Connect-PnPOnline -Url $SourceUrl -UseWebLogin;

$PageList = Get-PnPListItem -List "SitePages";
$Records = New-Object Collections.Generic.List[Item];
$ResultPath = Resolve-Path ".\Result";

if (-not (Test-Path $ResultPath)) {
    New-Item $ResultFolder -ItemType Directory;
}

$Index = 0;
foreach ($PageItem in $PageList) {
    $Name = $PageItem.FieldValues["FileLeafRef"];
    $Guid = $PageItem.FieldValues["GUID"];
    $TemplateFolderPath = Join-Path -Path $ResultPath -ChildPath $Guid;
    $TemplateFilePath = Join-Path -Path $TemplateFolderPath -ChildPath "Template.xml";

    New-Item -ItemType directory -Path $TemplateFolderPath | Out-Null;

    Write-Progress `
        -Activity "Exportando $($Name)..." `
        -Status "$($Index) de $($PageList.Count) páginas." `
        -PercentComplete (($Index / $PageList.Count) * 100);

    try {
        Export-PnPClientSidePage `
            -Identity $Name `
            -Out $TemplateFilePath `
            -PersistBrandingFiles `
            -ErrorAction Stop `
            -Force;
    }
    catch {
        Write-Error "Erro ao exportar página: $($Name)!";
    }

    $Time = Get-Date -Format "o";
    $Record = [Item]::new($TemplateFilePath, $Name, $Time, $Guid);
    $Records.Add($Record);
    $Index = $Index + 1;
}

$LogPath = Join-Path -Path $ResultPath -ChildPath "Log.csv";
$Records | Export-Csv -Path $LogPath -Force;