
# Exemple:
# set-location WhateverDirectoryUnderWhichYouWantToCreateANewMudBlazorSolution;
# New-MudBlazorSolution.ps1 -AppName 'Blazor Demo 03' -PageNames @('E Commerce', 'Settings') -Force;
param(
    [parameter(Mandatory=$true , Position = 0, HelpMessage = "The application name, to be placed at the main Navigation Menu at the left of the main page.", ValueFromPipeline=$false)]
    [string]$AppName,

    [parameter(Mandatory=$true , Position = 1, HelpMessage = "A list of Page Names, this script will create empty pages then insert entries for them at the Navigation Menu. You can use spaces, letters and numbers at the page names (but they must start with a letter), the script will remove the spaces for the file names, which will have proper Pascal name convention.", ValueFromPipeline=$false)]
    [string[]]$PageNames,

    [parameter(Mandatory=$false , Position = 2, HelpMessage = "Name of Key Vault that will receive the Application client secret. Mandatory if using AddAuthentication.", ValueFromPipeline=$false)]
    [string]$KeyVaultName,

    [parameter(Mandatory=$false, Position = 3, HelpMessage = "Optional, default is 5.0.8. The version for the MudBlazor package. See https://www.nuget.org/packages/MudBlazor for more details on current versions.", ValueFromPipeline=$false)]
    [string]$MudBlazorVersion = '5.0.8',

    [parameter(Mandatory=$false, HelpMessage = "Optional switch, if you use -AddAuthentication, this script will add authentication to the project. This is not fully functional, but can help if you *really* know what you are doing.", ValueFromPipeline=$false)]
    [switch]$AddAuthentication,

    [parameter(Mandatory=$false, HelpMessage = "Optional switch, if you use -Force, this script will not ask any confirmations.", ValueFromPipeline=$false)]
    [switch]$Force
)

function disableAllNugetSourcesExcept([string]$sourceName)
{
    $sourceList = nuget.exe sources list;
    $currentSourceConfiguration = @{};
    for ($i = 0; $i -lt $sourceList.Count; $i++)
    {
        $line = $sourceList[$i];
        if ($line -match '^[ \t]*[0-9]+\.[ \t]+(.*) \[(.*)\]')
        {
            $name  = $Matches[1];
            $state = $Matches[2];
            $currentSourceConfiguration[$name] = $state;
            $operation = if ($name -eq $sourceName) { 'Enable' } else { 'Disable'};

            nuget.exe sources $operation -Name $name;
        }
    }

    return $currentSourceConfiguration;
}

function restoreNugetSourcesState([Hashtable]$sourcesFromDisableAllNugetSourcesExcept)
{
    foreach ($name in ($sourcesFromDisableAllNugetSourcesExcept.Keys | select))
    {
        $operation = if ($sourcesFromDisableAllNugetSourcesExcept[$name] -match 'Enable' ) { 'Enable' } else { 'Disable'};
        nuget.exe sources $operation -Name $name;
    }
}

Write-Host "Creating MudBlazor solution." -ForegroundColor Green;

try 
{
    $context = Get-AzContext;
}
catch 
{
    if (0x80131501 -eq $_.Exception.HResult)
    {
        throw "Looks like Az.Accounts module is not installed. Please install it and try again. To install, run`nInstall-Module Az.Accounts -Force; Import-Module Az.Accounts -Force;";
    }
}

if ($null -eq $context)
{
    throw "Please log in to Azure, use Connect-AzAccount -UseDeviceAuthentication"
}

if (!$Force)
{
    Write-Host "The current Azure context has Account == [$($context.Account.Id)] and Subscription Name == [$($context.Subscription.Name)]" -ForegroundColor Yellow;
    Write-Host "Is this where you want to place the $AppName app?" -ForegroundColor Yellow;
    $yn = Read-Host "Yes/No (Y/N)";
    if ($yn -ne 'y')
    {
        throw "Cancelled.";
    }
}

if (![string]::IsNullOrEmpty($KeyVaultName) -and $AddAuthentication)
{
    throw "Please inform Key vault name if you want to add authentication.";
}

Write-Host "Checking if you have dotnet..." -ForegroundColor Green;
$dotnetVersion = Invoke-Expression 'dotnet --version' -ErrorAction SilentlyContinue 2>&1;
if ($null -eq $dotnetVersion)
{
    throw "Please install dotnet, from https://dotnet.microsoft.com/download";
}

if ([int]::Parse($dotnetVersion.Split('.')[0]) -lt 5)
{
    throw "Please install dotnet 5 or above, from https://dotnet.microsoft.com/download";
}

. (Join-Path $PSScriptRoot 'mudBlazorGenericFunctions.ps1');
$solutionName = getClassName $AppName; 

if ($AddAuthentication)
{
    Write-Host "Checking if you have some necessary tools (msidentity-app-sync, Az.Account)" -ForegroundColor Green;
    $dotnetToolList = [string]::Join('#', (dotnet tool list -g));
    if ($dotnetToolList -notmatch 'msidentity-app-sync')
    {
        [HashTable]$previousSourcesState = disableAllNugetSourcesExcept 'nuget.org';
        dotnet tool install -g msidentity-app-sync;
        restoreNugetSourcesState $previousSourcesState;
    }


    $appUri   = "https://$solutionName.onmicrosoft.com";
    $app      = Get-AzADApplication -IdentifierUri $appUri;
    Write-Host "Trying to create Azure App Registration under URI $appUri" -ForegroundColor Green;

    $tenantId = $context.Tenant.Id;

    if ($null -ne $app)
    {
        throw "Azure App with uri $appUri already exists, I will not create a new one. If you want to remove the current app, run`nRemove-AzAdApplication -ApplicationId $($app.ApplicationId.Guid) -Force";
    }

    $secretName = $solutionName + 'AppKey';
    $app        = New-AzADApplication -DisplayName $AppName -IdentifierUris $appUri;

    Write-Host "Creating secret, to be placed on Key Vault $KeyVaultName, on secret named $secretName." -ForegroundColor Green;
    $keyOpen    = (New-Guid).Guid;
    $key        = ConvertTo-SecureString -String $keyOpen -AsPlainText -Force;
    New-AzADAppCredential -ApplicationId $app.ApplicationId -Password $key -EndDate ((get-date).AddYears(1));
    Set-AzKeyVaultSecret -VaultName $KeyVaultName -Name $secretName -SecretValue $key;
}

Write-Host "Creating Blazor solution." -ForegroundColor Green;
$solutionDir  = Join-Path $PWD $solutionName;
dotnet new sln          -o $solutionDir --name $solutionName;

Write-Host "Creating Blazor project." -ForegroundColor Green;
Set-Location $solutionName;
$projectName  = "$solutionName.BlazorServer";
$projectPath  = Join-Path $PWD $projectName;
if ($AddAuthentication)
{
    dotnet new blazorserver -o $projectPath --name $projectName --auth SingleOrg --tenant-id $tenantId --calls-graph;
}
else 
{
    dotnet new blazorserver -o $projectPath --name $projectName;
}

Set-Location $projectPath;

. (Join-Path $PSScriptRoot 'mudify.ps1') -AppName $AppName -PageNames $PageNames -MudBlazorVersion $MudBlazorVersion;
Write-Host "Done mudifying $projectName project." -ForegroundColor Green;

if ($AddAuthentication)
{
    Write-Host "Adding MS Identity to project $projectName." -ForegroundColor Green;
    msidentity-app-sync --tenant-id $tenantId --client-id $app.ObjectId --client-secret $keyOpen;
}

Write-Host "Adding $projectName Blazor project to solution $solutionName." -ForegroundColor Green;
Set-Location ..
dotnet sln add "$projectName/$projectName.csproj"

Write-Host "Finding Visual Studio." -ForegroundColor Green;

if ($null -eq (get-module VSSetup)) { Install-Module -Name VSSetup -Force -AllowClobber; }

if ($null -eq (Get-VSSetupInstance))
{
    Write-Host "Visual Studio not installed." -ForegroundColor Magenta;
    return;
}

$maxVSVersion  = (Get-VSSetupInstance | ForEach-Object { $_.InstallationVersion} | Sort-Object -Descending)[0];
$vsInstance    =  Get-VSSetupInstance | Where-Object { $_.InstallationVersion -eq $maxVSVersion };
$vsInstallPath = $vsInstance.InstallationPath;
$devenvPath    = (Get-ChildItem -Filter 'devenv.*' -Path $vsInstallPath -File -Recurse).FullName | Select-Object -First 1;
Write-Host "Found Visual studio at $devenvPath." -ForegroundColor Green;
$workPath      = [System.IO.Path]::GetDirectoryName($devenvPath);

$solutionPath = Resolve-Path "$solutionName.sln";
Write-Host "Opening $solutionPath." -ForegroundColor Green;
Start-Process $devenvPath -WorkingDirectory $workPath -ArgumentList "$solutionPath";

Write-Host "Done creating $AppName solution"  -ForegroundColor Green;
