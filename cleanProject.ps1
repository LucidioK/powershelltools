function delDir($dirName)
{
    $fns = (Get-ChildItem -Filter $dirName -Directory -Recurse).FullName;
    foreach ($fn in $fns)
    {
        Remove-Item -Path $fn -Recurse -Force | Out-Null;
    };
}

$ErrorActionPreference = 'Stop';
pushd .

while ((Get-ChildItem  -Path '.' -Filter '*.sln') -eq $null -and $PWD.Path.Length -gt 3)
{
    Set-Location '..';
}

if ($PWD.Path.Length -gt 3)
{
    #git reset --hard;
    $ln = @(); 
    for ($1 = 0; $i -lt 2048; $i++) { $ln += 'n'; }
    try
    {
        $ln | git clean -fdx 2>&1 | Out-Null;
        delDir 'bin';
        delDir 'obj';
        delDir '.vs';
    }
    catch
    {
        Write-Host "`n`nMake sure you either close Visual Studio or close the solution (File / Close Solution). Aborting." -ForegroundColor Magenta;
        popd;
        return;
    }
}
else
{
    Write-Host 'Could not find solution file. Nothing was done.' -ForegroundColor Magenta;
}

popd