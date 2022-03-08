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
enum InjectionPosition { Before; After; VeryBeginning; VeryEnd; At; }

function getEncoding([string]$path)
{
  if ($PSVersiontable.PSEdition -eq 'Core')
  {
    [byte[]]$bytes = get-content -AsByteStream  -ReadCount 4 -TotalCount 4 -Path $Path;
  }
  else
  {
    [byte[]]$bytes = get-content -Encoding byte -ReadCount 4 -TotalCount 4 -Path $Path;
  }
  $encoding = 'ascii';
  if ( $bytes[0] -eq 0xef -and $bytes[1] -eq 0xbb -and $bytes[2] -eq 0xbf )
  { $encoding =  'utf8' }
  elseif ($bytes[0] -eq 0xfe -and $bytes[1] -eq 0xff)
  { $encoding =  'bigendianunicode' }
  elseif ($bytes[0] -eq 0xff -and $bytes[1] -eq 0xfe)
  { $encoding =  'unicode' }
  elseif ($bytes[0] -eq 0 -and $bytes[1] -eq 0 -and $bytes[2] -eq 0xfe -and $bytes[3] -eq 0xff)
  { $encoding =  'UTF32 Big-Endian' }
  elseif ($bytes[0] -eq 0xfe -and $bytes[1] -eq 0xff -and $bytes[2] -eq 0 -and $bytes[3] -eq 0)
  { $encoding =  'utf32' }
  elseif ($bytes[0] -eq 0x2b -and $bytes[1] -eq 0x2f -and $bytes[2] -eq 0x76 -and ($bytes[3] -eq 0x38 -or $bytes[3] -eq 0x39 -or $bytes[3] -eq 0x2b -or $bytes[3] -eq 0x2f) )
  { $encoding =  'utf7'}
  elseif ( $bytes[0] -eq 0xf7 -and $bytes[1] -eq 0x64 -and $bytes[2] -eq 0x4c )
  { $encoding =  'UTF-1' }
  elseif ($bytes[0] -eq 0xdd -and $bytes[1] -eq 0x73 -and $bytes[2] -eq 0x66 -and $bytes[3] -eq 0x73)
  { $encoding =  'UTF-EBCDIC' }
  elseif ( $bytes[0] -eq 0x0e -and $bytes[1] -eq 0xfe -and $bytes[2] -eq 0xff )
  { $encoding =  'SCSU' }
  elseif ( $bytes[0] -eq 0xfb -and $bytes[1] -eq 0xee -and $bytes[2] -eq 0x28 )
  { $encoding =  'BOCU-1' }
  elseif ($bytes[0] -eq 0x84 -and $bytes[1] -eq 0x31 -and $bytes[2] -eq 0x95 -and $bytes[3] -eq 0x33)
  { $encoding =  'GB-18030' }

  return $encoding;
}

function getFileLinesAsList
{
    param(
        [string]$FilePath
    )

    $lines = [System.Collections.Generic.List[string]]::new();
    Get-Content $filePath | ForEach-Object { $lines.Add($_); }

    return $lines;
}

function findFile
{
    param(
        [string]$Filter, 
        [bool]$FailIfNotFound = $true
    )

    $filePath = (Get-ChildItem -Filter $Filter -Recurse).FullName;
    
    if ([string]::IsNullOrEmpty($filePath) -and $FailIfNotFound)
    {
        throw "File $filePath not found under the current directory. Make sure you are in a Blazor project directory.";
    }
    
    return $filePath;
}

function deleteFileIfFound
{
    param(
        [string]$FilePath
    )

    if (![string]::IsNullOrEmpty($filePath))
    {
        Write-Host "  Deleting $FilePath." -ForegroundColor Green;
        Remove-Item -Path $FilePath;
    }
}

function findLinePosition
{
    param(
        [System.Collections.Generic.List[string]]$lines, 
        [InjectionPosition]$InjectionPosition, 
        [string]$RegexPattern
    )

    if ([string]::IsNullOrEmpty($regexPattern))
    {
        throw "Provide a regex if you want to use At, Before or After.";
    }

    $position = -1;
    for ($i = 0; $i -lt $lines.Count; $i++)
    {
        if ($lines[$i] -match $regexPattern)
        {
            $position = $i;
            if ($injectionPosition -eq [InjectionPosition]::After)
            {
                $position++;
            }

            break;
        }
    }

    if ($position -eq -1)
    {
        throw "Could not find $regexPattern on $filePath."
    }

    return $position;
}

function addLineToFile
{
    param(
        [string]$FilePath, 
        [string]$LineToAdd, 
        [InjectionPosition]$InjectionPosition = [InjectionPosition]::Before, 
        [string]$RegexPattern = $null
    )

    [System.Collections.Generic.List[string]]$lines = getFileLinesAsList $FilePath;
    $encoding = getEncoding $FilePath;

    if ($injectionPosition -eq [InjectionPosition]::VeryBeginning)
    {
        $positionToAdd = 0;
    }
    elseif ($injectionPosition -eq [InjectionPosition]::VeryEnd)
    {
        $positionToAdd = $lines.Count;
    }
    else
    {
        if (![string]::IsNullOrEmpty($regexPattern))
        {
            $positionToAdd = findLinePosition $lines $InjectionPosition $RegexPattern;
        }
        else
        {
            throw "Provide a regex if you want to use At, Before or After.";
        }
    }

    $lines.Insert($positionToAdd, $lineToAdd);
    $lines | Out-File $filePath -Encoding $encoding;
}

function removeLineFromFile
{
    param(
        [string]$FilePath, 
        [string]$RegexPattern
    )

    [System.Collections.Generic.List[string]]$lines = getFileLinesAsList $FilePath;
    $encoding = getEncoding $FilePath;
    $position = findLinePosition $lines ([InjectionPosition]::At) $RegexPattern;
    $lines.RemoveAt($position);
    $lines | Out-File $filePath -Encoding $encoding;
}

function getPageClassName
{
    param(
        [string]$PageName
    )

    $PageName = $PageName -replace ' +',' ';
    $PageName = $PageName.Trim();
    $sn = $PageName.Normalize([Text.NormalizationForm]::FormD);
    $sb = New-Object Text.StringBuilder;
 
    for ($i = 0; $i -lt $sn.Length; $i++) 
    {
        $c = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($sn[$i]);
        if($c -ne [Globalization.UnicodeCategory]::NonSpacingMark) 
        {
          [void]$sb.Append($sn[$i]);
        }
    }
    
    $sbn = $sb.ToString().Normalize([Text.NormalizationForm]::FormC);
    $className = "";
    $pleaseCapitalizeNextCharacter = $true;
    for ($i = 0; $i -lt $sbn.Length; $i++) 
    {
        $c = $sbn[$i];
        if ($c -eq ' ')
        {
            $pleaseCapitalizeNextCharacter = $true;
            continue;
        }
        
        if ($pleaseCapitalizeNextCharacter)
        {
            $c = [char]::ToUpperInvariant($c);
            $pleaseCapitalizeNextCharacter = $false;
        }

        if ($c -eq '?')
        {
            $c = '_';
        }
        
        $className += $c;        
    }

    return $className;
}

function addPage
{
    param(
        [string]$PageName,
        [string]$ProjectName
    )

    $className = getPageClassName $PageName;
    $pageRef   = "        <MudNavLink Href=""$className"" Icon=""@Icons.Material.Filled.Extension"">$PageName</MudNavLink>"

    addLineToFile -FilePath          (findFile 'NavMenu.razor')   `
                  -LineToAdd         $pageRef                     `
                  -InjectionPosition ([InjectionPosition]::After) `
                  -RegexPattern      'Icons.Material.Filled.Home';

    @"
@page "/$className"

@using System.Text.RegularExpressions
@using Microsoft.AspNetCore.Components.Rendering
@using System.Threading
@using System.Collections.Concurrent
@using $ProjectName.Data 
@inject IDataService DataService
<MudGrid>
    <MudItem xs="12">
       <MudText Typo="Typo.h6" Class="px-4 mt-4 mb-4">TBD: $PageName</MudText>
    </MudItem>
</MudGrid>
"@ | Out-File "Pages\$className.razor" -Encoding utf8;

    @"
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace $ProjectName.Pages
{
    public partial class $className 
    {
        public void OnGet()
        {
        }
    }
}
"@ | Out-File "Pages\$className.razor.cs" -Encoding utf8;
}

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
              -RegexPattern      '</head>|<component .*/>';

$scriptLineToAdd        = '<script src="_content/MudBlazor/MudBlazor.min.js"></script>';
addLineToFile -FilePath          $cssFontReferenceFilePath     `
              -LineToAdd         $scriptLineToAdd              `
              -InjectionPosition ([InjectionPosition]::After)  `
              -RegexPattern      '<body|<component .*/>';

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
