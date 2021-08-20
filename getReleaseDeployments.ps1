param(
    [parameter(Mandatory=$true , Position = 0)]
    [string]$PAT,
        
    [parameter(Mandatory=$true , Position = 1)]
    [string]$Organization, 

    [parameter(Mandatory=$true , Position = 2)]
    [string]$Project, 

    [parameter(Mandatory=$true , Position = 3)]
    [int]$DefinitionId,

    [parameter(Mandatory=$false, Position = 4)]
    [int]$Top = 200
)

function getHdr()
{
    $basicAuth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$PAT"));

    $header = @{
        "authorization" = "Basic $basicAuth";
        "accept"        = "application/json;api-version=6.1-preview.1;excludeUrls=true"
    };

    return $header;
}

function getWeb($url)
{
    $r = Invoke-WebRequest -Method Get -Uri $url -Headers (getHdr);

    return $r;
}

function getRest($url)
{
    $r = Invoke-RestMethod -Method Get -Uri $url -Headers (getHdr);
    if ($null -ne $r -and $null -ne $r.value)
    {
        $r = $r.value;
    }

    return $r;
}


$releaseDeployments = @();
$continuationToken = $null;

$topLength = $Top.ToString().Length;
do
{
    if ($null -eq $continuationToken)
    {
        $url = "https://vsrm.dev.azure.com/$Organization/$Project/_apis/Release/deployments?definitionId=$DefinitionId&api-version=6.0";
        Write-Host "Reading deployment from release definition $($DefinitionId) from Organization $Organization, Project $Project, initial read." -ForegroundColor Green; 
    }
    else
    {
        $url = "https://vsrm.dev.azure.com/$Organization/$Project/_apis/Release/deployments?definitionId=$DefinitionId&continuationToken=$continuationToken&api-version=6.0";
        Write-Host "Reading deployment from release definition $($DefinitionId) from Organization $Organization, Project $Project, continue reading." -ForegroundColor Green; 
    }

    Write-Host "Reading from $url." -ForegroundColor Green; 
    $result = getWeb $url;

    if ($result.StatusCode -eq 200)
    {
        $continuationToken   = $result.Headers['x-ms-continuationtoken'];
        $resultContent       = ($result.Content | ConvertFrom-Json).value;

        foreach ($r in $resultContent)
        {
            Write-Host " Reading $(($releaseDeployments.Count+1).ToString().PadLeft($topLength))/$($Top) release $($r.release.name) from release definition $($r.releaseDefinition.name)" -ForegroundColor Green; 
            $release             = getRest $r.release.url;
            $release.createdBy   = $release.createdBy.uniqueName;
            $release.createdFor  = $release.createdFor.uniqueName;
            $release.modifiedBy  = $release.modifiedBy.uniqueName;
            $releaseDeployments += $release;
            if ($releaseDeployments.Count -ge $Top)
            {
                break;
            }
        }
    }
}
while ($result.StatusCode -eq 200 -and $releaseDeployments.Count -lt $Top);

Write-Host "Done.`n`n" -ForegroundColor Green; 
$releaseDeployments = $releaseDeployments | Select-Object -First $Top;

return $releaseDeployments;
