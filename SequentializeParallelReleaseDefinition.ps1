<#
.SYNOPSIS

  Converts a Release Definition from parallel execution to sequential.

.DESCRIPTION

  Gets a given Release Definition, then places all its stages in the order defined by Rank, saving it as a new Release Definition with the same name with suffix " (Sequential)".

.PARAMETER <PAT>
PAT (Personal Access Token) for ADO organization to manipulate Release Definitions, 
to be obtained at https://ORGANIZATIONNAME.visualstudio.com/_usersSettings/tokens.

.PARAMETER <OrganizationName>
ADO organization name (aka ADO Instance name).

For example, when you open your Release Definition on your browser, if your URL was as shown below, the Organization Name would be AAAAA.

https://AAAAA.visualstudio.com/XXXXX/_releaseDefinition?definitionId=123456&_a=environments-editor-preview


.PARAMETER <ProjectName>
ADO organization name (aka ADO Instance name).

For example, when you open your Release Definition on your browser, if your URL was as shown below, the Project Name would be XXXXX.

https://AAAAA.visualstudio.com/XXXXX/_releaseDefinition?definitionId=123456&_a=environments-editor-preview

.PARAMETER <ReleaseDefinitionId>
Release definition id.

For example, when you open your Release Definition on your browser, if your URL was as shown below, the Release Definition ID would be 123456.

https://AAAAA.visualstudio.com/XXXXX/_releaseDefinition?definitionId=123456&_a=environments-editor-preview


.PARAMETER <FolderPath>
Path to save the Release Definition into.
Optional, default is \

.OUTPUTS
The new Release Definition as an object.

#>
param(
    [parameter(Mandatory=$true , Position = 0, HelpMessage = 'PAT (Personal Access Token) for ADO organization to manipulate Release Definitions, to be obtained at https://ORGANIZATIONNAME.visualstudio.com/_usersSettings/tokens.')]
    [string]$PAT,
    [parameter(Mandatory=$true , Position = 1, HelpMessage = 'ADO organization name (aka ADO Instance name).')]
    [string]$OrganizationName,
    [parameter(Mandatory=$true , Position = 2, HelpMessage = 'ADO Project name.')]
    [string]$ProjectName, 
    [parameter(Mandatory=$true , Position = 3, HelpMessage = 'Release definition id.')]
    [int]$ReleaseDefinitionId,
    [parameter(Mandatory=$false , Position = 4, HelpMessage = 'Release definition folder path. Optional, default is to save in the root.')]
    [string]$FolderPath = '\\')

function getReleaseDefinition($PAT, $OrganizationName, $projectName, $ReleaseDefinitionId)
{
    Invoke-RestMethod `
        -URI "https://vsrm.dev.azure.com/$OrganizationName/$ProjectName/_apis/Release/definitions/$ReleaseDefinitionId" `
        -Headers @{ "authorization" = "Basic $([Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$PAT")))";};
}

function newReleaseDefinition($PAT, $OrganizationName, $projectName, $releaseDefinition)
{
    $releaseDefinition.id = $null; # Must remove the id, since it will be added by the POST below.
    Invoke-RestMethod `
        -Method POST `
        -URI "https://vsrm.dev.azure.com/$OrganizationName/$ProjectName/_apis/Release/definitions?api-version=6.0" `
        -Headers @{ "authorization" = "Basic $([Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$PAT")))"; 'Content-Type' = 'application/json'} `
        -Body ($releaseDefinition | ConvertTo-Json -Depth 32);
}

$rd = getReleaseDefinition $PAT $OrganizationName $ProjectName $ReleaseDefinitionId;
$rankToEnvironmentName = @{};
for ($i = 0; $i -lt $rd.environments.Count; $i++)
{
    $rankToEnvironmentName[$rd.environments[$i].rank] = $rd.environments[$i].name;
}

$environmentsToSequentialize = $rd.environments | Where-Object { $null -ne $_.conditions }  | Sort-Object -Property rank;

for ($i = 1; $i -lt $environmentsToSequentialize.Count; $i++)
{
    $previousRank = $environmentsToSequentialize[$i - 1].rank;
    $previousEnvironmentName = $rankToEnvironmentName[$previousRank];
    $environmentsToSequentialize[$i].conditions = @( @{ name = $previousEnvironmentName; conditionType = 'environmentState'; value = '4'});
}

$rd.name = $rd.name + ' (Sequential)';
$rd.path = $FolderPath;
$newRD = newReleaseDefinition $PAT $OrganizationName $ProjectName $rd;

return $newRD;
