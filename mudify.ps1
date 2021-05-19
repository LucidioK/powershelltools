#
# Changes Blazor project to use MudBlazor, as documented at https://mudblazor.com/getting-started/installation
# Just to to the directory of your .csproj and run mudify.ps1
# Save your code before running this script, no guarantees!
#
param(
    [parameter(Mandatory=$true , Position = 0, HelpMessage = "The application name, to be placed at the main Navigation Menu at the left of the main page.", ValueFromPipeline=$false)]
    [string]$AppName,

    [parameter(Mandatory=$true , Position = 1, HelpMessage = "A list of Page Names, this script will create empty pages then insert entries for them at the Navigation Menu. You can use spaces, letters and numbers at the page names (but they must start with a letter), the script will remove the spaces for the file names, which will have proper Pascal name convention.", ValueFromPipeline=$false)]
    [string[]]$PageNames,

    [parameter(Mandatory=$false, Position = 2, HelpMessage = "Optional, default is 5.0.8. The version for the MudBlazor package. See https://www.nuget.org/packages/MudBlazor for more details on current versions.", ValueFromPipeline=$false)]
    [string]$MudBlazorVersion = '5.0.8'
)

Write-Host 'Starting mudify.' -ForegroundColor Green;
$ErrorActionPreference = 'Stop';
. (Join-Path $PSScriptRoot 'mudBlazorGenericFunctions.ps1');

###############################################################################
Write-Host '1. Install the package' -ForegroundColor Green;
$linesToAdd = @"
  <ItemGroup>
    <PackageReference Include="MudBlazor" Version="$MudBlazorVersion" />
  </ItemGroup>
"@;

addLineToFile -FilePath          (findFile '*.csproj')         `
              -LineToAdd         $linesToAdd                   `
              -InjectionPosition ([InjectionPosition]::Before) `
              -RegexPattern      '</Project>';

###############################################################################
Write-Host '2. Add Imports' -ForegroundColor Green;
addLineToFile -FilePath          (findFile '_Imports.razor')   `
              -LineToAdd         '@using MudBlazor'            `
              -InjectionPosition ([InjectionPosition]::VeryEnd);

###############################################################################
Write-Host '3. Add CSS & Font references' -ForegroundColor Green;
$cssFontReferenceFilePath = findFile 'index.html' $false;
$thisIsABlazorWebAssemblyProject = $true;
if ($cssFontReferenceFilePath -eq $null)
{
    $cssFontReferenceFilePath = findFile '_Host.cshtml';
    $thisIsABlazorWebAssemblyProject = $false;
}

$cssFontLinesToAdd      = @"
<link href="https://fonts.googleapis.com/css?family=Roboto:300,400,500,700&display=swap" rel="stylesheet" />
<link href="_content/MudBlazor/MudBlazor.min.css" rel="stylesheet" />
"@;

addLineToFile -FilePath          $cssFontReferenceFilePath     `
              -LineToAdd         $cssFontLinesToAdd            `
              -InjectionPosition ([InjectionPosition]::Before) `
              -RegexPattern      '</head>';

$scriptLineToAdd        = '<script src="_content/MudBlazor/MudBlazor.min.js"></script>';
addLineToFile -FilePath          $cssFontReferenceFilePath     `
              -LineToAdd         $scriptLineToAdd              `
              -InjectionPosition ([InjectionPosition]::After)  `
              -RegexPattern      '<body';

###############################################################################
Write-Host '4. Register Services' -ForegroundColor Green;
if ($thisIsABlazorWebAssemblyProject)
{
    Write-Host '    This is a Blazor WebAssembly project.'  -ForegroundColor Green;

    $filePath = (findFile 'Startup.cs');

    addLineToFile -FilePath          $filePath                                `
                  -LineToAdd         '    builder.Services.AddMudServices();' `
                  -InjectionPosition ([InjectionPosition]::After)             `
                  -RegexPattern      'await builder.Build\(\).RunAsync\(\)';
}
else
{
    Write-Host '    This is a Blazor Server project.'  -ForegroundColor Green;

    $filePath = (findFile 'Startup.cs');

    addLineToFile -FilePath          $filePath                                `
                  -LineToAdd         '            services.AddMudServices();' `
                  -InjectionPosition ([InjectionPosition]::After)             `
                  -RegexPattern      'services.AddServerSideBlazor';
}

addLineToFile -FilePath          $filePath                                    `
              -LineToAdd         'using MudBlazor.Services;'                  `
              -InjectionPosition ([InjectionPosition]::Before)                `
              -RegexPattern      '^namespace ';

addLineToFile -FilePath          $filePath                                    `
              -LineToAdd         '            services.AddSingleton<IDataService, DataService>();' `
              -InjectionPosition ([InjectionPosition]::After)                 `
              -RegexPattern      'services.AddMudServices';

removeLineFromFile -FilePath     $filePath                                    `
                   -RegexPattern 'WeatherForecastService';

###############################################################################
Write-Host '5. Add Components' -ForegroundColor Green;
$filePath = (findFile 'App.razor');
addLineToFile -FilePath          $filePath                                                       `
              -LineToAdd         '<MudThemeProvider/><MudDialogProvider/><MudSnackbarProvider/>' `
              -InjectionPosition ([InjectionPosition]::VeryBeginning);

###############################################################################
Write-Host '6. Removing WeatherForecast and Counter, that comes with created Blazor project' -ForegroundColor Green;
deleteFileIfFound (findFile 'WeatherForecast.cs'        $false);
deleteFileIfFound (findFile 'WeatherForecastService.cs' $false);
deleteFileIfFound (findFile 'Counter.razor'             $false);
deleteFileIfFound (findFile 'FetchData.razor'           $false);


###############################################################################
Write-Host '7. Starting New Navigation Page' -ForegroundColor Green;
@"
<MudPaper Width="200px" >
    <MudNavMenu>
        <MudText Typo="Typo.h6" Class="px-4 mt-4 mb-4">$AppName</MudText>
        <MudDivider />
        <MudNavLink Href=""          Icon="@Icons.Material.Filled.Home"     >Home    </MudNavLink>
    </MudNavMenu>
</MudPaper>
"@ | Out-File -FilePath (findFile 'NavMenu.razor') -Encoding utf8;

addLineToFile -FilePath          (findFile 'MainLayout.razor')                                    `
              -LineToAdd         '            <!-- TBD: Update about -->' `
              -InjectionPosition ([InjectionPosition]::Before)                 `
              -RegexPattern      'docs.microsoft.com';

###############################################################################
Write-Host '8. Adding your pages' -ForegroundColor Green;

$projectName = [System.IO.Path]::GetFileNameWithoutExtension((findFile '*.csproj'));
for ($i = $PageNames.Count - 1; $i -ge 0; $i--)
{
    $pageName = $PageNames[$i];
    Write-Host "   Adding $pageName"  -ForegroundColor Green;
    addPage $pageName $projectName;
}


###############################################################################
Write-Host '9. Creating empty data service' -ForegroundColor Green;
$csCode = @"
namespace $projectName.Data
{
    public interface IDataService
    {
        // TBD: add interface members
    }
}
"@ | Out-File 'Data\IDataService.cs' -Encoding utf8;

@"
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Serialization;
using System.Threading;

namespace $projectName.Data
{
    public class DataService : IDataService
    {
        // TBD: add class members
    }
}
"@ | Out-File 'Data\DataService.cs' -Encoding utf8;

###############################################################################
Write-Host '10. Simplifying Index.razor' -ForegroundColor Green;

@"
@page "/"
<MudPaper Class="pa-16 ma-2" Outlined="true" Square="true"  Elevation="3">
    <MudText Typo="Typo.h6" Class="px-4 mt-4 mb-4">$AppName</MudText>
    <MudText Typo="Typo.h6" Class="px-4 mt-4 mb-4">TBD: Add content here.</MudText>
</MudPaper>
"@ | out-file 'Pages\Index.razor';

Write-Host 'Done.' -ForegroundColor Green;
