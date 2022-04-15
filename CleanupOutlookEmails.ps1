
param(
    [parameter(Mandatory=$true , Position = 1)]
    [string]$EmailAccount, # Lucidio.Kuhn@microsoft.com
    [parameter(Mandatory=$true , Position = 2)]
    [string]$FolderName, <# 'Inbox' #>    
    [parameter(Mandatory=$false , Position = 3)]
    [PSCustomObject[]]$EmailItems = $null

    )

function getConversationIdsToBePacked([PSCustomObject[]]$EmailItems)
{
    $conversationIds = ($EmailItems).ConversationID;
    $countByConversationId = @{};
    $conversationIds  | Select-Object -Unique | Select-Object | ForEach-Object { $countByConversationId.Add($_, 0); }
    foreach ($conversationId in $conversationIds)
    {
        $countByConversationId[$conversationId]++;
    }

    $conversationIdsToBePacked = @();

    foreach ($conversationId in ($countByConversationId.Keys | Select-Object))
    {
        if ($countByConversationId[$conversationId] -gt 1)
        {
            $conversationIdsToBePacked += $conversationId;
        }
    }

    return $conversationIdsToBePacked;
}

if ($null -eq $EmailItems)
{
    $EmailItems = getEmailList.ps1 -emailAccount $EmailAccount -folderName $FolderName; 
}


$conversationIdsToBePacked = getConversationIdsToBePacked $EmailItems;
$entryIdsToDelete = @();
foreach ($conversationIdToBePacked in $conversationIdsToBePacked)
{
    $entryIdsToDelete += ($EmailItems | Where-Object ConversationID -eq $conversationIdToBePacked | Sort-Object -Property ReceivedTime -Descending | Select-Object -Skip 1).EntryID;
}

DeleteEmailsFromOutlook.ps1 -EntryIds $entryIdsToDelete -EmailAccount $EmailAccount -FolderName $FolderName;

