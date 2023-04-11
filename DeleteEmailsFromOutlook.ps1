param(
    [parameter(Mandatory=$true , Position = 0)]
    [string[]]$EntryIds, 
    
    [parameter(Mandatory=$true , Position = 1)]
    [string]$EmailAccount, # Lucidio.Kuhn@microsoft.com
    [parameter(Mandatory=$true , Position = 2)]
    [string]$FolderName <# 'Inbox' #>)
if (!$env:Path.Contains($global:PSScriptRoot)) { $env:Path += ";$($global:PSScriptRoot)"}
$namespace     = GetOutlookNamespace.ps1;
$accountFolder = GetOutlookSubFolder.ps1 -Folders $namespace.Folders     -Name $emailAccount;
$emailFolder   = GetOutlookSubFolder.ps1 -Folders $accountFolder.Folders -Name $folderName;
$hs = [System.Collections.Generic.HashSet[string]]::new($EntryIds);
$numberdeleted = 0;
$count         = $emailFolder.Items.Count; 
for ($i = $count; $i -gt 0; $i--) 
{
    $j = $count + 1 - $i;
    Write-Progress -Activity "Deleting items $j/$count, $numberdeleted deleted" -PercentComplete ($j*100.0/$count);
    $item      = $emailFolder.Items($i); 
    [string]$entryId   = $item.EntryID;
    if ($hs.Contains($entryId))
    {
        Write-Host "Deleting email '$($item.Subject)', received $($item.ReceivedTime)"
        $item.UnRead = $false;
        $item.Save();
        $item.Delete();        
        $numberdeleted++;
    }
}

Get-Process OUTLOOK -ErrorAction SilentlyContinue | ForEach-Object { $_.Kill(); };

Write-Host "\n\nDone." -ForegroundColor Green;
