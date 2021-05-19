param(
    [parameter(Mandatory=$false, HelpMessage = "Use AsJob if you want to run SetPriority repeatedly, every 1 minute.", ValueFromPipeline=$false)]
    [switch]$AsJob,
    [string[]]$IdlePriorityProcesses = @(        
        'AlertusDesktop*',       
        'armsvc',       
        'Code.exe',         
        'EDPCleanup*',           
        'FileCoAuth*',           
        'IntuitUpdateService',   
        'MicrosoftSearchInBing', 
        'msiexec*', 
        'OneDrive',              
        'PerfWatson2',           
        'QualysAgent',           
        'RdrLeakDiag',           
        'Search*',         
        'spoolsv*',              
        'splwow*',               
        'sql*',                  
        'WmiPrvSE',              
        'TelemetryHost',         
        'TiWorker*'              
        ),
    [string[]]$BelowNormalPriorityProcesses = @(
        'CcmExec*',              
        'firefox*',
        'MSOID*',                
        'OfficeClickToRun',      
        'Nxt*',                  
        'policyHost',            
        'powershell*',           
        'posh',       
        'samsungdex*',           
        'SenseCE*',              
        'SenseNDR*',  
        'ServiceFabric*',           
        'TGitCache'),
    [string[]]$NormalPriorityProcesses = @(
        'chrome',
        'msedge*',
        'procexp*',
        'Teams*'
    ),
    [string[]]$ProcessesToKill = @(
        'downloader2*',
        'realdownload*',
        'skype'
    )
    )


function global:SetPriority
{
    function GetProcessListByName([string]$name)
    {
        [System.Collections.Generic.List[System.Diagnostics.Process]]$processes = [System.Collections.Generic.List[System.Diagnostics.Process]]::new();
        $ps = Get-Process -Name $processName -ErrorAction SilentlyContinue;
        if ($ps -ne $null)
        {
            if ($ps.GetType().Name -eq 'Process')
            {
                $processes.Add($ps);
            }
            else
            {
                $ps | foreach { $processes.Add($_); }
            }
        }

        return $processes;
    }

    function GetProcessTree(
            [int]$processId, 
            [object[]]$listUntilNow = $null, 
            [object[]]$processHierarchy = $null)
    {
        Start-Sleep -Milliseconds 1;

        if ($processHierarchy -eq $null)
        {
            $processHierarchy = Get-WmiObject win32_process | select ProcessId, ParentProcessId;
        }

        if ($listUntilNow -eq $null)
        {
            $process = get-process -Id $processId;
            $listUntilNow = @($process);
        }

        $childProcessIds = ($processHierarchy | ? parentprocessid -EQ $processId).ProcessId;
        foreach ($childProcessId in $childProcessIds)
        {
            [System.Collections.Generic.HashSet[int]]$idsUntilNow = [System.Collections.Generic.HashSet[int]]::new(($listUntilNow).Id);
            if ($idsUntilNow -eq $null -or $idsUntilNow.Count -eq 0 -or !($idsUntilNow.Contains($childProcessId)))
            {
                $listUntilNow += (get-process -Id $childProcessId); 
                $listUntilNow = GetProcessTree $childProcessId $listUntilNow $processHierarchy;
            }
        }

        return $listUntilNow;
    }

    function SetPriorityForThisProcess([System.Diagnostics.Process]$process,[string]$priorityClass)
    {
        $processTree = GetProcessTree -processId $process.Id;
        foreach ($p in $processTree)
        {
            Write-Host "$priorityClass : $($p.Name) " -ForegroundColor Green;
            Start-Sleep -Milliseconds 1;
            $p.PriorityClass = $priorityClass;
        }
    }

    function killProcesses([string[]]$processNames)
    {
        foreach ($processName in $processNames)
        {
            Write-Host "Killing $processname" -ForegroundColor Green;
            Stop-Process -Name  $processname -ErrorAction SilentlyContinue;
        }
    }

    function setPriorities([string[]]$processNames, [string]$priorityClass)
    {
        foreach ($processName in $processNames)
        {
            try
            {
                GetProcessListByName $processName | 
	            ForEach-Object { 
                    SetPriorityForThisProcess -process $_ -priorityClass $priorityClass;
                }
            }
            catch
            {
                write-host "Oops $($_.Exception.Message)" -ForegroundColor Yellow;
            }
        }
    }

    $thisProcess = Get-Process -Id $PID;
    try
    {
        $thisProcess.PriorityClass = 'Idle';
        setPriorities $IdlePriorityProcesses 'Idle';
        setPriorities $BelowNormalPriorityProcesses 'BelowNormal';
        $thisProcess.PriorityClass = 'Idle';
        setPriorities $NormalPriorityProcesses 'Normal';
        $thisProcess.PriorityClass = 'Idle';
	    killProcesses $ProcessesToKill;
    }
    finally
    {
        $thisProcess.PriorityClass = 'BelowNormal';
    }
}

if ($AsJob)
{
    $block = { while ($true) { SetPriority; Start-Sleep -Seconds 60; } };
    Start-Job -ScriptBlock $block;
}
else
{
    SetPriority;
}
