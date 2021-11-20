param(
    [parameter(Mandatory=$true , Position = 0)]
    [string]$ChatOrChannelName, 

    [parameter(Mandatory=$false , Position = 1)]
    [switch]$Channel,

    [parameter(Mandatory=$false , Position = 2)]
    [string]$OutputFilePath = ($ChatOrChannelName + ".txt"), 

    [parameter(Mandatory=$false , Position = 3)]
    [int]$MaximumNumberOfLines = (8 * 1024)    
)

function CopyPreviousLineFromTeams([int]$processId)
{
    $wshell.AppActivate($processId) | out-null;
    Start-Sleep -Milliseconds 100;
    $ocb = '{'; $ccb = '}';
    [System.Windows.Forms.SendKeys]::SendWait("$($ocb)UP$ccb")
    Start-Sleep -Milliseconds 100;
    $text = $null;
    for ($i = 0; $i -lt 4 -and [string]::IsNullOrEmpty($text); $i++)
    {
        [System.Windows.Forms.SendKeys]::SendWait('^(c)');
        Start-Sleep -Milliseconds (($i+1)*100);
        $text = Get-Clipboard;
    }

    if (!([string]::IsNullOrEmpty($text)))
    {
        # the text comes in the format UserName TimeStamp Message. However, the timestamp is not separated from
        # the UserName and Message, so I am adding spaces here.
        $timeStamp = [DateTime]::Parse((extractWithRegex $text "($timeStampPattern)"));
        $timeStamp = $timeStamp.ToString('yyyy/MM/dd HH:mm')
        $userName  = extractWithRegex $text "(.*?)$timeStampPattern";
        $message   = extractWithRegex $text ".*?$timeStampPattern(.*)";
        $text = "$timeStamp $($userName): $message";
    }
    return $text;  
}

function extractWithRegex([string]$str, [string]$patternWithOneGroupMarker)
{
    if ($str -match $patternWithOneGroupMarker)
    {
        return $matches[1];
    }
    return $null;
}

function getMessageTimestamp([string]$msg)
{
    $ts        = extractwithregex $msg '([0-9]{4}/[0-9]{2}/[0-9]{2} [0-9]{2}:[0-9]{2})';
    $year      = extractwithregex $ts '([0-9]{4})/[0-9]{2}/[0-9]{2} [0-9]{2}:[0-9]{2}';
    $month     = extractwithregex $ts '[0-9]{4}/([0-9]{2})/[0-9]{2} [0-9]{2}:[0-9]{2}';
    $day       = extractwithregex $ts '[0-9]{4}/[0-9]{2}/([0-9]{2}) [0-9]{2}:[0-9]{2}';
    $hour      = extractwithregex $ts '[0-9]{4}/[0-9]{2}/[0-9]{2} ([0-9]{2}):[0-9]{2}';
    $minute    = extractwithregex $ts '[0-9]{4}/[0-9]{2}/[0-9]{2} [0-9]{2}:([0-9]{2})';
    $timeStamp = [DateTime]::new($year, $month, $day, $hour, $minute, 0);

    return $timeStamp;
}

$timeStampPattern = '[0-9]{1,2}/[0-9]{1,2}/[0-9]{1,4} [0-9]{1,2}:[0-9]{1,2} [AP]M';
$FileName = [System.IO.Path]::GetFileName($OutputFilePath);
# Convert all weird on the file name characters to _
$FileName = $FileName  -replace '[^A-Za-z0-1\.]','_';
$Directory = [System.IO.Path]::GetDirectoryName($OutputFilePath);
if ([string]::IsNullOrEmpty($Directory))
{
    $OutputFilePath = $FileName;
}
else 
{
    $OutputFilePath = Join-Path $Directory $FileName;    
}

if (Test-Path $OutputFilePath)
{
    $previousLines = Get-Content $OutputFilePath;
    $latestDate    = ($previousLines | ForEach-Object { getMessageTimestamp $_; } | Measure-Object -Maximum).Maximum;
}
else 
{
    $previousLines = @();
    $latestDate    = [DateTime]::MaxValue;
}

if ($Channel) { $cc = 'channel' } else { $cc = 'chat' };
Write-Host "Important!" -ForegroundColor Yellow;
Write-Host "Before starting this script, do this:" -ForegroundColor Yellow;
Write-Host "1. Open Teams." -ForegroundColor Yellow;
Write-Host "2. Open the $cc named '$ChatOrChannelName'"  -ForegroundColor Yellow;
Write-Host "3. Click on the last message." -ForegroundColor Yellow;
Write-Host "";
Write-Host "Press Enter to continue:";
Read-Host;

[string]$ChatOrChannelNameClean = "";
for ($i = 0; $i -lt $ChatOrChannelName.Length; $i++)
{
    $c = $ChatOrChannelName[$i];
    if ($c -notmatch '[A-Za-z0-9 ]') { $c = ('\' + $c); }
    $ChatOrChannelNameClean += $c;
}

$processId =  (Get-Process *teams*  | Where-Object { $_.MainWindowTitle -match $ChatOrChannelNameClean} ).id;
if ($null -eq $processId)
{
    throw "Could not find a teams process with main window title '$ChatOrChannelName'";
}

Add-Type -AssemblyName System.Windows.Forms;
$wshell = New-Object -ComObject WScript.Shell;
$previousMessage = "<NONE>";
$sameMessageCounter = 0;
$lines = @();
for ($i = 0; $i -lt $MaximumNumberOfLines; $i++) 
{
    [string]$messageText = CopyPreviousLineFromTeams $processId;
    Write-Progress -Activity $messageText.PadRight(64).Substring(0, 64) -PercentComplete 0;
    $timeStamp = getMessageTimestamp $messageText;
    if ($timeStamp -le $latestDate)
    {
        break;
    }

    if ($messageText -eq $previousMessage)
    {
        $sameMessageCounter++;
    }
    else 
    {
        $sameMessageCounter = 0;
    }

    if ($sameMessageCounter -gt 4 -or $null -eq $messageText)
    {
        break;
    }
    if ($sameMessageCounter -eq 0)
    {
        $lines = @($messageText) + $lines;
    }
    $previousMessage = $messageText;
}

$lines = $previousLines + $lines;
$lines | Out-File $OutputFilePath;

Write-Host "$i messages captured into $(Resolve-Path $OutputFilePath)" -ForegroundColor Green;
