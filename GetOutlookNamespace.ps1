if (!$env:Path.Contains($global:PSScriptRoot)) { $env:Path += ";$($global:PSScriptRoot)"}
Get-Process OUTLOOK -ErrorAction SilentlyContinue | ForEach-Object { $_.Kill(); };

if ($null -eq ('Microsoft.Office.Interop.Outlook.olDefaultFolders' -as [type]))
{
    $officeDirectories = (Get-ChildItem -Path $env:ProgramFiles -Filter '*Office*' -Directory | Sort-Object -Descending -Property CreationTime).FullName;
    foreach ($officeDirectory in $officeDirectories)
    {
        $outlookInteropPath = (Get-ChildItem -Path $officeDirectory -Filter Microsoft.Office.Interop.Outlook.dll -File -Recurse).Fullname;
        if ($null -ne $outlookInteropPath)
        {
            break;
        }
    }

    if ($null -eq $outlookInteropPath)
    {
        throw "Could not find Microsoft.Office.Interop.Outlook.dll. Make sure Office is installed, with Outlook, under folder $($env:ProgramFiles).";
    }
    Add-type -assembly $outlookInteropPath | out-null;
}

$outlook = new-object -comobject outlook.application;
$namespace = $outlook.GetNameSpace('MAPI');
Write-Host "
|                                                                                               |
|                                                                                               |
|                                                                                               |
| IMPORTANT:                                                                                    |
| IMPORTANT:                                                                                    |
| IMPORTANT:                                                                                    |
| If the script seems to be hanging, look at the taskbar, see if you have an Outlook icon there.|
| It might be asking for your permission to muck with the emails.                               |
|                                                                                               |" -ForegroundColor Black -BackgroundColor Magenta;
return $namespace;