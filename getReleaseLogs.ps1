param(
    [parameter(Mandatory=$true , Position = 0)]
    [string]$PAT,    

    [parameter(Mandatory=$true , Position = 1)]
    [PSCustomObject[]]$ReleaseList, # from getReleaseDeployments.ps1

    [parameter(Mandatory=$false , Position = 2)]
    [bool]$OnlyFromFailures = $true
)

function filterByStatus([PSCustomObject[]]$l)
{
    if ($OnlyFromFailures)
    {
        return ($l | Where-Object { $_.status -ne 'succeeded'});
    }

    return $l;
}

function getWeb([string]$url)
{
    if ([string]::IsNullOrEmpty($url))
    {
        return $null;
    }

    $r = $null;
    for ($i = 0; $i -lt 4; $i++)
    {
        try 
        {
            $r = Invoke-WebRequest -Method Get -Uri $url -Headers $script:header;
            break;
        }
        catch 
        {
            if ($_.Exception.Message -match 'did not properly respond after a period of time') 
            {
                Write-Host "Host did not respond for $url, wait 2 seconds and retry ($($i+1)/4)";
                Start-Sleep -Seconds 2;
            }
        }
    }

    if ($null -eq $r)
    {
        return $null;
    }

    return $r.content;
}

$basicAuth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$PAT"));

$script:header = @{
    "authorization" = "Basic $basicAuth";
    "accept"        = "application/json;api-version=6.1-preview.1;excludeUrls=true"
};
$logs = @();

foreach ($release in $ReleaseList)
{
    foreach ($environment in (filterByStatus $release.environments))
    {
        foreach ($deployStep in (filterByStatus $environment.deploySteps))
        {
            foreach ($releaseDeployPhase in (filterByStatus $deployStep.releaseDeployPhases))
            {
                foreach ($deploymentJob in (filterByStatus $releaseDeployPhase.deploymentJobs))
                {
                    foreach ($task in (filterByStatus $deploymentJob.tasks))
                    {
                        $item = [PSCustomObject]@{
                            ReleaseId                 = $release.id;
                            ReleaseName               = $release.name;
                            ReleaseStatus             = $release.Status;
                            ReleaseCreatedOn          = $release.createdOn;
                            EnvironmentId             = $environment.id;
                            EnvironmentName           = $environment.name;
                            DeployStepId              = $deployStep.id;
                            DeployStepDeploymentId    = $deployStep.deploymentId;
                            DeployStepAttempt         = $deployStep.attempt;
                            DeployStepReason          = $deployStep.reason;
                            DeployStepStatus          = $deployStep.status;
                            DeployStepOperationStatus = $deployStep.operationStatus;
                            DeployPhaseId             = $releaseDeployPhase.id;
                            DeployPhaseName           = $releaseDeployPhase.name;
                            DeployPhaseRank           = $releaseDeployPhase.rank;
                            DeployPhaseStatus         = $releaseDeployPhase.status;
                            DeploymentJobId           = $deploymentJob.job.id;
                            DeploymentJobName         = $deploymentJob.job.name;
                            DeploymentJobStatus       = $deploymentJob.job.status;
                            DeploymentJobLogUrl       = $deploymentJob.job.logUrl;
                            DeploymentJobLog          = (getWeb $deploymentJob.job.logUrl);
                            TaskId                    = $task.id;
                            TaskTimelineRecordId      = $task.timelineRecordId;
                            TaskName                  = $task.name;
                            TaskDateStarted           = $task.dateStarted;
                            TaskDateEnded             = $task.dateEnded;
                            TaskStartTime             = $task.startTime;
                            TaskFinishTime            = $task.finishTime;
                            TaskStatus                = $task.status;
                            TaskRank                  = $task.rank;
                            TaskIssues                = $task.issues;
                            TaskLogUrl                = $task.logUrl;
                            TaskDefId                 = $task.task.id;
                            TaskDefName               = $task.task.name;
                            TaskDefVersion            = $task.task.version;
                            TaskLog                   = (getWeb $task.logUrl);
                        };

                        if (!$OnlyFromFailures `
                            -or !([string]::IsNullOrEmpty($item.DeploymentJobLog)) `
                            -or !([string]::IsNullOrEmpty($item.TaskLog)))
                        {
                            $logs += $item;
                        }
                    }
                }
            }
        }
    }
}

return $logs;
