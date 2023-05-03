###
# Description:
#   Converts a resume (CV) from yaml to DOCX.
#   - You must have Microsoft(r) Word installed.
#   - The yaml file follows the schema described below.
#
# yaml schema:
# ---
# name: <NAME>
# description:
#   - <DESCRIPTION_1>
# email: <EMAIL>
# phone: <PHONE>
# linkedIn: <LINKEDIN_PAGE>
# address: <PHYSICAL_ADDRESS>
# citizenship:
#   - <CITIZENSHIP_1>
#   - <CITIZENSHIP_2>
# needSponsorship: <NEED_SPONSORSHIP_YES_NO>
# skills:
#   - <SKILL_1>
# programmingLanguages:
#   - <PROGRAMMING_LANGUAGE_1>
# techStacks:
#   - teckStack:
#     cloud: <CLOUD_1_AZURE_AWS_GCC>
#     stack:
#       - <STACK_ITEM_1>
# professionalExperiences:
#   - professionalExperience:
#     company: <COMPANY_N>
#     title: <TITLE_N>
#     startingMonth: <STARTING_MONTH_N>
#     endingMonth: <STARTING_MONTH_N>
#     details:
#       - <DETAIL_N_1>
#     links:
#       - <LINK_N_1>
# education:
#   - detail:
#     degree: <DEGREE_1>
#     area: <AREA_1>
#     university: <UNIVERSITY_1>
#     country: <EDUCATION_COUNTRY_1>
# 
# Parameters:
#   yaml_file_path: path of YAML file to convert to docx.
#   Force: (Optional) if you inform this switch, this script will kill all Word processes without asking you if want to do so. 
#          By default, this script will stop it execution and ask whether you want to kill all Word processes running on this computer.
# Example:
#     yaml_resume_to_word.ps1 c:\temp\cvlk.yaml
#     yaml_resume_to_word.ps1 c:\temp\cvlk.yaml -Force
#
# Output:
#   It will create a .docx file, with the extension appended to the file name.
#   In the example above, the created file would be c:\temp\cvlk.yaml.docx
#
param(
    [parameter(mandatory=$true , position = 0)]
    [string]$yaml_file_path,
    [parameter(mandatory=$false , position = 1)]
    [switch]$Force)

function ReadYN([string]$prompt)
{
    $prompt = $prompt + ' y/n:';
    [string]$answer = '_';
    Write-Host;
    do
    {
        Write-Host $prompt -ForegroundColor Green;
        $answer = (Read-Host).ToUpperInvariant();
    } 
    while ($answer -ne 'Y' -and $answer -ne 'N');

    return ($answer -eq 'Y');
}
function KillAllWordProcessesIfNeeded([bool]$forceKill)
{
    $wordProcesses = Get-Process 'WinWord' -ErrorAction SilentlyContinue;
    if ($forceKill -and  $null -ne $wordProcesses)
    {
        Write-Host "There are Word processes that will be killed, but it seems you know what you are doing..." -ForegroundColor Yellow;
    }
    
    if (!$forceKill -and  $null -ne $wordProcesses)
    {
        $forceKill = ReadYN("There are Word processes running. Would you like to kill all these processes without saving your work? ");
    }

    if ($forceKill)
    {
        Get-Process 'WinWord' -ErrorAction SilentlyContinue | ForEach-Object { $_.Kill() };
    }
    elseif ($null -ne $wordProcesses) 
    {
        Write-Host;
        Write-Host;
        throw "Please save your work, close all Word process and retry.";
    }
}

if ($null -eq ('WordBase' -as [Type]))
{
    Import-Module (Join-Path $PSScriptRoot 'WordBase.ps1');
}

if ($null -eq (Get-Command 'ConvertTo-Yaml' -ErrorAction 'SilentlyContinue'))
{
    Install-Module 'powershell-yaml' -Force;
    Import-Module 'powershell-yaml';
}

Write-Host "Starting conversion." -ForegroundColor Green;

KillAllWordProcessesIfNeeded $Force;

Write-Host " Starting Word." -ForegroundColor Green;
try
{
    $w = [WordBase]::new();
}
catch 
{
    Write-Host;
    Write-Host "Could not open Word. Try closing this PowerShell session, open another one then try again." -ForegroundColor Magenta;
    exit 1;
}

Write-Host "  Adjusting text styles." -ForegroundColor Green;
$w.ChangeStyle('Heading 1', 'Trebuchet MS', 16, $false, $true, 0, 0)
$w.ChangeStyle('Heading 2', 'Trebuchet MS', 14, $false, $true, 0, 0)
$w.ChangeStyle('Heading 3', 'Trebuchet MS', 12, $false, $true, 0, 0)
$w.ChangeStyle('Normal', 'Trebuchet MS',    11, $false, $false, 0, 0)

Write-Host "  Retrieving and converting the YAML file." -ForegroundColor Green;
$resume = Get-Content $yaml_file_path | convertfrom-yaml;

Write-Host " Writing header."  -ForegroundColor Green;
Write-Host "  Writing person's name as large text at the top of the document, centralized."  -ForegroundColor Green;
$w.StyledText($resume.name, 'Heading 1', $w.wdParagraphAlignment.wdAlignParagraphCenter);

Write-Host "  Writing phone, email and linkedin link, centralized." -ForegroundColor Green;
$w.StyledText($resume.phone + ' - ' + $resume.email + ' - ' + $resume.linkedIn, 'Normal', $w.wdParagraphAlignment.wdAlignParagraphCenter);

Write-Host "  Writing the description items as bullet list." -ForegroundColor Green;
$w.SelectStyle('Normal');
$w.BulletList($resume.description);


Write-Host " Writing skills items as a normal paragraph." -ForegroundColor Green;
$w.StyledText('Skills', 'Heading 2', $w.wdParagraphAlignment.wdAlignParagraphLeft)
$w.StyledText([string]::Join(' ', $resume.skills), 'Normal', $w.wdParagraphAlignment.wdAlignParagraphJustify)
$w.StyledText([string]::Join(' ', $resume.programmingLanguages), 'Normal', $w.wdParagraphAlignment.wdAlignParagraphJustify)

Write-Host "  Writing tech stacks." -ForegroundColor Green;
$w.StyledText('Tech Stacks', 'Heading 2', $w.wdParagraphAlignment.wdAlignParagraphLeft)
foreach ($techStack in ($resume.techStacks).techStack)
{
    Write-Host "   Writing tech stack $($techStack.cloud) as normal paragraph, after a heading with the name of the cloud." -ForegroundColor Green;
    $w.StyledText("On $($techStack.cloud):", 'Heading 3', $w.wdParagraphAlignment.wdAlignParagraphLeft)
    $w.StyledText([string]::Join(' ', $techStack.stack), 'Normal', $w.wdParagraphAlignment.wdAlignParagraphJustify)
}

Write-Host " Writing Professional Experience" -ForegroundColor Green;
$w.StyledText('Professional Experience', 'Heading 2', $w.wdParagraphAlignment.wdAlignParagraphLeft)
foreach ($pe in ($resume.professionalExperiences).professionalExperience)
{
    Write-Host "  Writing professional experience $($pe.title) - $($pe.company)" -ForegroundColor Green;
    $w.StyledText("$($pe.title) - $($pe.company) - $($pe.startingMonth) - $($pe.endingMonth)", 'Heading 3', $w.wdParagraphAlignment.wdAlignParagraphLeft)
    $w.SelectStyle('Normal');
    $w.BulletList($pe.details);
    $w.BulletList($pe.links);
}

Write-Host " Writing education items." -ForegroundColor Green;
$w.StyledText('Education', 'Heading 2', $w.wdParagraphAlignment.wdAlignParagraphLeft);

Write-Host "  Writing education items as bullet points." -ForegroundColor Green;
$educationList = ($resume.education).detail | ForEach-Object { "$($_.degree) - $($_.area) - $($_.university)" } ;
$w.SelectStyle('Normal');
$w.BulletList($educationList);

Write-Host " Saving document $($yaml_file_path + ".docx")" -ForegroundColor Green;
$w.Save($yaml_file_path + ".docx");

Write-Host " Closing word." -ForegroundColor Green;
$w.Close();
Write-Host "Done." -ForegroundColor Green;
