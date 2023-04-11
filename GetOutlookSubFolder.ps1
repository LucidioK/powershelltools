
param(
    [parameter(Mandatory=$true , Position = 0)]
    [object]$Folders, 
    
    [parameter(Mandatory=$true , Position = 0)]
    [string]$Name) 

if (!$env:Path.Contains($global:PSScriptRoot)) { $env:Path += ";$($global:PSScriptRoot)"}

$entryId = $null;
$count = $Folders.Count;
for ($i = 1; $null -eq $entryId -and $i -le $count; $i++)
{
    $item = $Folders.Item($i);
    [string]$foldername = $item.Name;
    if ($foldername -eq $Name)
    {
        [string]$entryId = $item.EntryId;
    }
}

if ($null -eq $entryId)
{
    throw "Folder $Name not found";
}
$namespace = $Folders.Application.GetNamespace('MAPI');
$folder = $namespace.GetFolderFromID($entryId);
return $folder;
