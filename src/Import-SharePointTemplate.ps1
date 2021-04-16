using module ".\Item.psm1";

[CmdletBinding()]
param (
    # URL of the destination site.
    [Parameter(Mandatory)]
    [string]
    $DestinationUrl,

    # Convert to team site
    [Parameter(Mandatory = $false)]
    [switch]
    $ConvertToTeamSite = $false
)

function Invoke-Template {
    param (
        [Parameter(Mandatory)]
        [string]
        $Name,

        [Parameter(Mandatory)]
        [int]
        $Total,

        [Parameter(Mandatory)]
        [string]
        $TemplatePath,

        [Parameter(Mandatory)]
        [int]
        $Index
    )

    $Count = 0;
    do {
        $Count = $Count + 1;
        try {
            Write-Progress `
                -Activity "Importando $($Name). Tentativa $($Count) de 3" `
                -Status "$($Index) de $($Total) páginas" `
                -PercentComplete (($Index / $Total) * 100);
            Apply-PnPProvisioningTemplate -Path $TemplatePath -ClearNavigation;
            return;
        }
        catch {
            Connect-PnPOnline -Url "https://spo-global.kpmg.com/sites/BR-SMSASAPMovement" -UseWebLogin;
            Start-Sleep -Seconds 30;
        }
    } while ($Count -lt 3)

    $Failure = "Falha na importação da página $($Name).";
    throw $Failure;
}


Connect-PnPOnline -Url $DestinationUrl -UseWebLogin;

$Index = 0;
$Templates = Import-Csv -Path ".\Result\Log.csv";
$Total = $Templates.Count;

foreach ($Template in $Templates) {
    $Name = $Template.Name;
    $TemplatePath = $Template.Path;
    
    if ($ConvertToTeamSite) {
        $FolderPath = Split-Path -Path $TemplatePath
        $AltTemplatePath = Join-Path -Path $FolderPath -ChildPath "AltTemplate.xml";
        Copy-Item $TemplatePath -Destination $AltTemplatePath;
        (Get-Content $AltTemplatePath).replace('OneColumnFullWidth', 'OneColumn') | Set-Content $AltTemplatePath;
        $TemplatePath = $AltTemplatePath;
    }
    
    try {
        Invoke-Template -Name $Name -Total $Total -TemplatePath $TemplatePath -Index $Index -ErrorAction Stop;
    }
    catch {
        Write-Error "Fim da execução. Encerrada no item $($Name), na posição $($Index).";
        break;
    }
    $Index = $Index + 1;
}