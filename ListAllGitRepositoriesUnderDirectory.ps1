param(
    [Parameter(Mandatory=$true)]
    [string]$Directory
)
function extractWithRegex([string]$str, [string]$patternWithOneGroupMarker)
{
    if ($str -match $patternWithOneGroupMarker)
    {
        return $matches[1];
    }
    return $null;
}

$dirs = (Get-ChildItem ".git" -Recurse -Depth 2 -Attributes Hidden -Path $Directory).FullName | ForEach-Object { $_.Replace('\.git','') };
$curdir = (Get-Location).Path;
foreach ($dir in $dirs)
{
    Set-Location $dir;
    $fetchUrl = ((git remote -v | Where-Object { $_ -match '\(fetch\)' }) -replace 'origin[ \t]+','') -replace ' .*','';
    $project = extractwithregex $fetchUrl '([^/]+)/_git';
    $repo = extractwithregex $fetchUrl '/_git/([^/]+)';
    Write-Host "https://skype.visualstudio.com/$project/_git/$repo";
}

Set-Location $curdir;