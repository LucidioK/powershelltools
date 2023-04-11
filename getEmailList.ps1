param ([string]$emailAccount = 'Lucidio.Kuhn@microsoft.com',
       [string]$folderName = 'Inbox')

if (!$env:Path.Contains($global:PSScriptRoot)) { $env:Path += ";$($global:PSScriptRoot)"}
function removeRes([string]$s)
{
    do
    {
        $initialSize = $s.Length;
        $s = $s -replace '^Re: *','';
        $s = $s -replace '^Fw: *','';
        $finalSize = $s.Length;
    } 
    until ($initialSize -eq $finalSize);

    return $s;
}

function substr([string]$s, [int]$l)
{
    if ($s.Length -lt $l)
    {
        $s = $s.PadRight($l - $s.Length);
    }
    else 
    {
        $s = $s.Substring(0, $l);    
    }

    return $s;
}

function getEmailList([string]$emailAccount, [string]$folderName = 'Inbox')
{
    $namespace     = GetOutlookNamespace.ps1;
    $accountFolder = GetOutlookSubFolder.ps1 -Folders $namespace.Folders     -Name $emailAccount;
    $emailFolder   = GetOutlookSubFolder.ps1 -Folders $accountFolder.Folders -Name $folderName;
    $mailItems     = @(); 
    $ct            = $emailFolder.Items.Count; 
    for ($i = 1; $i -le $ct; $i++) 
    { 
        $item      = $emailFolder.Items($i); 
        $subject   = removeRes $item.Subject;
        $s48       = substr $subject 48;
        Write-Progress -Activity "Reading from $folderName [$s48]" -PercentComplete ($i*100.0/$ct);
        $mailItems += [PSCustomObject]@{
            Subject            = $subject;
            EntryID            = $item.EntryID;
            ReceivedTime       = $item.ReceivedTime;
            AttachmentCount    = $item.Attachments.Count;
            Categories         = $item.Categories;
            SenderEmailAddress = $item.SenderEmailAddress;  
            ConversationID     = $item.ConversationID;  
        };
    }
    
    Get-Process OUTLOOK  -ErrorAction SilentlyContinue | ForEach-Object { $_.Kill(); };

    return $mailItems | Sort-Object -Property ReceivedTime -Descending;
}

$l = getEmailList $emailAccount $folderName;
return $l;
