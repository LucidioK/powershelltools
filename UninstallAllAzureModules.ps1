$moduleFolders = $env:PSModulePath.Split(';');
$moduleInfoPaths    = @{};
$installedModules = @();
function a2dl([Hashtable]$d,[string]$l,[object]$o)
{
    if (!($d.ContainsKey($l))) { $d.Add($l, @()); }
    $d[$l] += $o;
}

function getModuleId($moduleName, $moduleVersion)
{
    if ($moduleVersion.GetType() -eq [string]) { $moduleVersion = [System.Version]::new($moduleVersion); }
    return "$moduleName|$($moduleVersion.Major).$($moduleVersion.Minor).$($moduleVersion.Build)";
}

function getModuleNameFromModuleId([string]$moduleId)    { $moduleId.Split('|')[0] };
function getModuleVersionFromModuleId([string]$moduleId) { $moduleId.Split('|')[1] };

function pathLevel([string]$path) { ($path.Replace('/','\').ToCharArray() | Where-Object { $_ -eq '\'}).Count; }
function removeDirectory([string]$path)
{
    if (!(Test-Path $path))
    {
        return;
    }

    Get-ChildItem -Path $path -Recurse -File -Force | Remove-Item -Force;
    $dirs = (Get-ChildItem -Path $path -Recurse -Directory).FullName | ForEach-Object { "$((pathLevel $_).ToString("0000"))$_" } | Sort-Object -Descending;
    foreach ($dir in $dirs)
    {
        removeDirectory $dir.Substring(4);
    }

    while ($true)
    {
        $result = (Remove-Item -Path $path -Recurse -Force) *>&1;
        if ($null -eq $result)
        {
            break;
        }

        Start-Sleep -Milliseconds 200;
    }
}

function getModuleRootDirectories([string]$nameOfModuleToUninstall)
{
    $psds = @();
    $moduleIds = $moduleInfoPaths.Keys | Select-Object | Where-Object { $_.StartsWith("$nameOfModuleToUninstall|")};
    foreach ($moduleId in $moduleIds)
    {
        $psds += $moduleInfoPaths[$moduleId];
    }

    $roots = @();
    foreach ($psd in $psds)
    {
        $dir = [System.IO.Path]::GetDirectoryName($psd);
        $dir = [System.IO.Path]::GetDirectoryName($dir);
        $roots += $dir;
    }
    $roots = $roots | Select-Object -Unique | Sort-Object;

    return $roots;
}

foreach ($moduleFolder in $moduleFolders)
{
    $psdPaths = (Get-ChildItem -Path $moduleFolder -Filter 'Az*.psd1' -File -Recurse).FullName;
    foreach ($psdPath in $psdPaths)
    {
        $installedModules += [System.IO.Path]::GetFileNameWithoutExtension($psdPath);
    }
}

$installedModules = $installedModules | Select-Object -Unique | Sort-Object;

do 
{
    Write-Host "`nChecking module dependencies (this might take a while)..." -ForegroundColor Green;
    $moduleInfos        = @{};
    $moduleInfoPaths    = @{};
    $moduleDependencies = @{};
    $dependantModules   = @{};
    foreach ($moduleFolder in $moduleFolders)
    {
        $psdPaths = (Get-ChildItem -Path $moduleFolder -Filter 'Az*.psd1' -File -Recurse).FullName;
        foreach ($psdPath in $psdPaths)
        {
            $moduleInfo   = Import-PowerShellDataFile $psdPath;
            $moduleName   = [System.IO.Path]::GetFileNameWithoutExtension($psdPath);
            $moduleId     = getModuleId $moduleName $moduleInfo['ModuleVersion'];
            a2dl $moduleInfos     $moduleId $moduleInfo;
            a2dl $moduleInfoPaths $moduleId $psdPath   ;
            foreach ($reqMod in $moduleInfo['RequiredModules'])
            {
                $version      = if ($reqMod.ContainsKey('RequiredVersion')) { $reqMod['RequiredVersion'] } else { $reqMod['ModuleVersion'] }
                $dependencyId = getModuleId $reqMod['ModuleName'] $version;
                a2dl $moduleDependencies $moduleId $dependencyId;
                a2dl $dependantModules   $dependencyId $moduleId;
            }
            
        }
    }
    
    $allModuleNames                 = $moduleInfos.Keys      | Select-Object | ForEach-Object { getModuleNameFromModuleId $_; };
    $namesOfModulesWithDependencies = $dependantModules.Keys | Select-Object | ForEach-Object { getModuleNameFromModuleId $_; };
    $namesOfModulesWithNoDependencies = $allModuleNames | Where-Object { !$namesOfModulesWithDependencies.Contains($_) } | Select-Object -Unique | Sort-Object;
    
    foreach ($nameOfModuleToUninstall in $namesOfModulesWithNoDependencies)
    {
        Write-Host "Uninstalling $nameOfModuleToUninstall" -ForegroundColor Green;
        
        $moduleIdsToUninstall = $moduleInfos.Keys | Select-Object | Where-Object { $_.StartsWith("$nameOfModuleToUninstall|")};
        foreach ($moduleIdToUninstall in $moduleIdsToUninstall)
        {
            $moduleInfo = $moduleInfos[$moduleIdToUninstall];
            if ($moduleInfo.GetType() -ne [System.Management.Automation.PSModuleInfo])
            {
                $moduleInfo = [System.Management.Automation.PSModuleInfo]::new($moduleInfos[$moduleIdToUninstall]);
            }
            
            $version = getModuleVersionFromModuleId $moduleIdToUninstall;
            Write-Host " Removing module version $version" -ForegroundColor Green;

            Remove-Module -ModuleInfo $moduleInfo -Force; # -ErrorAction SilentlyContinue;
        }
    
        Write-Host " Uninstalling module $nameOfModuleToUninstall" -ForegroundColor Green;
        Uninstall-Module -Name $nameOfModuleToUninstall -AllVersions -Force;

        $directoriesToRemove = getModuleRootDirectories $nameOfModuleToUninstall;
        foreach ($directoryToRemove in $directoriesToRemove)
        {
            Write-Host " Removing $directoryToRemove" -ForegroundColor Green;
            removeDirectory $directoryToRemove;
        }
    }
} 
until ($dependantModules.Count -eq 0);

return $installedModules;