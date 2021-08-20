param(
    [Parameter(Mandatory=$true)]
    [string]$AssemblyPath
)

if (!(Test-Path $AssemblyPath))
{
    throw "$AssemblyPath not found.";
}

@('gacutil.exe','sn.exe','nuget.exe') | 
    ForEach-Object `
    {
        if ($null -eq (where.exe $_))
        {
            throw "Add the directory where $_ is into the PATH variable.";
        }
    }


function extractWithRegex([string]$str, [string]$patternWithOneGroupMarker)
{
    if ($str -match $patternWithOneGroupMarker)
    {
        return $matches[1];
    }
    return $null;
}

function getVersion([string]$v)
{
    $v = extractwithregex $v "^([0-9\.]+)";
    return ([System.Version]::new($v));
}

function getCandidates([string]$path, [string]$filePath)
{
    Get-ChildItem -Path $path -Filter $filePath -Recurse -File | 
        ForEach-Object { [PSCustomObject]@{ Path = $_.versioninfo.FileName; Version = (getVersion $_.versioninfo.ProductVersion)  } };
}

function getPublicKeyToken([string]$filePath)
{
    extractwithregex (sn.exe -T $filePath)  'Public key token is ([0-9a-z]+)';
}

function areVersionsMiminallyEqual([System.Version]$v1, [System.Version]$v2, [int]$level)
{
    switch ($level)
    {
        4 { return $v1.Major -eq $v2.Major -and $v1.Minor -eq $v2.Minor -and $v1.Build -eq $v2.Build -and $v1.Revision -eq $v2.Revision; }
        3 { return $v1.Major -eq $v2.Major -and $v1.Minor -eq $v2.Minor -and $v1.Build -eq $v2.Build; }
        2 { return $v1.Major -eq $v2.Major -and $v1.Minor -eq $v2.Minor; }
        1 { return $v1.Major -eq $v2.Major; }
    }

    return $false;
}

function tryLoadACandidateFromList([string[]]$filePathCandidates)
{
    foreach ($candidate in $filePathCandidates)
    {
      try {
        [System.Reflection.Assembly]::LoadFrom($candidate) | Out-Null;
        Write-Host "    Loaded $filePathCandidates successfully" -ForegroundColor Green;
        return $true;
      }
      catch { 
        if ($_.Exception.Message -match 'Assembly with same name is already loaded')
        {
          return $true;
        }
      }
    }

    return $false;
}

Write-Host "Loading $AssemblyPath" -ForegroundColor Green;
$asm = [Reflection.Assembly]::LoadFrom($AssemblyPath);
$gacDir = Join-Path $env:windir 'Microsoft.NET\assembly\GAC_MSIL';
Write-Host " GAC Directory $gacDir" -ForegroundColor Green;
$nugetDir = (nuget.exe locals global-packages -list).Replace('global-packages: ','');
Write-Host " Nuget Directory $nugetDir" -ForegroundColor Green;
$triedNugetInstall = $false;
$isDotNetCore = $PSVersionTable.PSEdition -eq 'Core';
$errout = $null;
do 
{
    Write-Host " Trying to retrieve types, to see if assembly is fine." -ForegroundColor Green;
    $types = invoke-expression '$asm.Modules[0].GetTypes()' -ErrorVariable errout -ErrorAction SilentlyContinue 2>&1 3>&1 | out-null;
    if ($null -ne $errout -and $errout.Count -gt 0)
    {
        $missingAssemblyName = (extractwithregex $errout[0].Exception.Message "assembly '(.*?),") ;
        $missingAssemblyFileName = $missingAssemblyName + '.dll';
        $missingAssemblyVersion = getVersion (extractwithregex $errout[0].Exception.Message "Version=([0-9\.]+)");
        $missingAssemblyPublicKeyToken = extractwithregex $errout[0].Exception.Message "PublicKeyToken=([0-9a-z]+)";
        Write-Host "  Referenced assembly $missingAssemblyName $missingAssemblyVersion $missingAssemblyPublicKeyToken not found..." -ForegroundColor Green;
        Write-Host "   Is it in the GAC?" -ForegroundColor Green -NoNewline;
        $isInGAC = $null -ne ((gacutil.exe /l $missingAssemblyName /nologo) -match $missingAssemblyName);
        $candidates = @();
        if ($isInGAC)
        {
            Write-Host " Yes!" -ForegroundColor Green;
            $candidates += getCandidates $gacDir $missingAssemblyFileName;
        }
        else 
        {
            Write-Host " No, no problem." -ForegroundColor Green;
        }

        Write-Host "   Searching on nuget cache... " -ForegroundColor Green -NoNewline;
        $candidates += getCandidates $nugetDir $missingAssemblyFileName;
        if ($isDotNetCore)
        {
            $candidates = $candidates | Where-Object { $_.Path  -match '\\net[0-9]+\.[0-9]+\\|\\netstandard'};
        }
        else 
        {
            $candidates = $candidates | Where-Object { $_.Path  -match '\\net[0-9]+\\|\\netstandard'};
        }
        Write-Host " Total $($candidates.Count) candidate files..." -ForegroundColor Green;

        $candidatesFilePathsWithCorrectVersion = @();
        for ($i = 4; $i -gt 0 -and $candidatesFilePathsWithCorrectVersion.Count -eq 0; $i--)
        {
            $candidatesFilePathsWithCorrectVersion = $candidates | Where-Object { areVersionsMiminallyEqual $_.Version $missingAssemblyVersion $i } | ForEach-Object { $_.Path };
        }

        Write-Host "   $($candidatesFilePathsWithCorrectVersion.Count) candidates with version $missingAssemblyVersion" -ForegroundColor Green;
        $candidatesFilePathsWithCorrectVersionAndPublicToken = $candidatesFilePathsWithCorrectVersion | Where-Object { (getPublicKeyToken $_) -eq $missingAssemblyPublicKeyToken };
        Write-Host "   $($candidatesFilePathsWithCorrectVersionAndPublicToken.Count) candidates with public token $missingAssemblyPublicKeyToken" -ForegroundColor Green;
        if ($candidatesFilePathsWithCorrectVersionAndPublicToken.Count -eq 0)
        {
            Write-Host "   Could not find assembly with public token key $missingAssemblyPublicKeyToken, will try only the ones with similar versions..." -ForegroundColor Yellow;
            $candidatesFilePathsWithCorrectVersionAndPublicToken = $candidatesFilePathsWithCorrectVersion;
        }

        if ($candidatesFilePathsWithCorrectVersionAndPublicToken.Count -eq 0)
        {
            if ($triedNugetInstall)
            {
                throw "Cannot find assembly $missingAssemblyName $missingAssemblyVersion $missingAssemblyPublicKeyToken on GAC nor nuget";
            }
            else 
            {
                Write-Host "   Trying to install nuget package $missingAssemblyName, then retrying." -ForegroundColor Green;
                nuget install $missingAssemblyName -Version $missingAssemblyVersion.ToString();
                $triedNugetInstall = $true;                    
            }
        }
        else 
        {
            tryLoadACandidateFromList $candidatesFilePathsWithCorrectVersionAndPublicToken;
            $triedNugetInstall = $false;
        }
    }
} 
until ($null -eq $errout -or $errout.Count -eq 0);

$types = $asm.GetTypes() | Where-Object { $_.IsPublic } | ForEach-Object { [PSCustomObject]@{
    Name = $_.Name;
    FullName = $_.FullName;
    BaseType = $_.BaseType.FullName;
    ImplementedInterfaces = ($_.ImplementedInterfaces).FullName;
    Fields = $_.DeclaredFields | ForEach-Object { [PSCustomObject]@{ Name = $_.Name; Type = $_.FieldType.Fullname }};
    Properties = $_.DeclaredProperties | ForEach-Object { [PSCustomObject]@{ Name = $_.Name; Type = $_.FieldType.Fullname }};
}};

return $types;