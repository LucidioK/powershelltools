enum InjectionPosition 
{ 
    Before; 
    After; 
    VeryBeginning; 
    VeryEnd; 
    At; 
}

function getEncoding([string]$path)
{
  [byte[]]$bytes = [System.IO.File]::ReadAllBytes($Path);
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

function getClassName
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

    $className = getClassName $PageName;
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
