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
# Example:
#     yaml_resume_to_word.ps1 c:\temp\cvlk.yaml
#
# Output:
#   It will create a .docx file, with the extension appended to the file name.
#   In the example above, the created file would be c:\temp\cvlk.yaml.docx
#
param(
    [parameter(mandatory=$true , position = 0)]
    [string]$yaml_file_path)

if ($null -eq ('WordBase' -as [Type]))
{
    Import-Module (Join-Path $PSScriptRoot 'WordBase.ps1');
}
if ($null -eq (Get-Command 'ConvertTo-Yaml' -ErrorAction 'SilentlyContinue'))
{
    Install-Module 'powershell-yaml' -Force;
    Import-Module 'powershell-yaml';
}
Get-Process 'WinWord' -ErrorAction SilentlyContinue | ForEach-Object { $_.Kill() };
$w = [WordBase]::new();
$w.ChangeStyle('Heading 1', 'Trebuchet MS', 16, $false, $true, 0, 0)
$w.ChangeStyle('Heading 2', 'Trebuchet MS', 14, $false, $true, 0, 0)
$w.ChangeStyle('Heading 3', 'Trebuchet MS', 12, $false, $true, 0, 0)
$w.ChangeStyle('Normal', 'Trebuchet MS',    11, $false, $false, 0, 0)

$resume = Get-Content $yaml_file_path | convertfrom-yaml;
$w.StyledText($resume.name, 'Heading 1', $w.wdParagraphAlignment.wdAlignParagraphCenter);

$w.StyledText($resume.phone + ' - ' + $resume.email + ' - ' + $resume.linkedIn, 'Normal', $w.wdParagraphAlignment.wdAlignParagraphCenter);
$w.StyledText([string]::Join(' ', $resume.description), 'Normal', $w.wdParagraphAlignment.wdAlignParagraphJustify)

$w.StyledText('Skills', 'Heading 2', $w.wdParagraphAlignment.wdAlignParagraphLeft)
$w.StyledText([string]::Join(' ', $resume.skills), 'Normal', $w.wdParagraphAlignment.wdAlignParagraphJustify)
$w.StyledText([string]::Join(' ', $resume.programmingLanguages), 'Normal', $w.wdParagraphAlignment.wdAlignParagraphJustify)

$w.StyledText('Tech Stacks', 'Heading 2', $w.wdParagraphAlignment.wdAlignParagraphLeft)
foreach ($techStack in ($resume.techStacks).techStack)
{
    $w.StyledText("On $($techStack.cloud):", 'Heading 3', $w.wdParagraphAlignment.wdAlignParagraphLeft)
    $w.StyledText([string]::Join(' ', $techStack.stack), 'Normal', $w.wdParagraphAlignment.wdAlignParagraphJustify)
}

$w.StyledText('Professional Experience', 'Heading 2', $w.wdParagraphAlignment.wdAlignParagraphLeft)
foreach ($pe in ($resume.professionalExperiences).professionalExperience)
{
    $w.StyledText("$($pe.title) - $($pe.company) - $($pe.startingMonth) - $($pe.endingMonth)", 'Heading 3', $w.wdParagraphAlignment.wdAlignParagraphLeft)
    $w.SelectStyle('Normal');
    $w.BulletList($pe.details);
    $w.BulletList($pe.links);
}

$w.StyledText('Education', 'Heading 2', $w.wdParagraphAlignment.wdAlignParagraphLeft);
foreach ($ed in ($resume.education).detail)
{
    $w.StyledText("$($ed.degree) - $($ed.area) - $($ed.university)", 'Normal', $w.wdParagraphAlignment.wdAlignParagraphLeft)
}

$w.Save($yaml_file_path + ".docx");
$w.Close();
