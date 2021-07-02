# powershelltools

A bunch of Powershell tools:

## Visio Tools

### ThreatModelToVisio.ps1

Converts a Microsoft Threat Modeling Tool diagram to Visio.

For instance, suppose you have a Threat Model in file `C:\temp\ThreatModel.tm7` like this:
![image](https://user-images.githubusercontent.com/1053501/124334102-eae6f700-db4a-11eb-8f8e-f459e0ff2765.png)

Notice that the threat model tab name is 'XYZ', in the example.

If you run this:

```
ThreatModelToVisio.ps1 -ThreatModelFilePath 'c:\temp\ThreatModel.tm7' -DiagramName 'XYZ';
```

The Visio diagram will be like this:
![image](https://user-images.githubusercontent.com/1053501/124334224-47e2ad00-db4b-11eb-8514-34b912d611ab.png)



### VisioBase.ps1

An all-purpose class to draw Visio diagrams with Powershell.

## Blazor Tools

### cleanProject.ps1

From time to time, a Visual Studio project gets in a dirty state and Blazor either cannot build or builds and misbehaves in unexplicable ways.

This script cleans up all necessary left-overs from the current path:
- All `bin` folders.
- All `obj` folders.
- All `.vs` folders.
- Plus, at the end, it will run `git clean -fdx`.

Be careful before using this script, since it will remove all files that are not under source control.

You should exit Visual Studio before running this script.

### mudBlazorGenericFunctions.ps1
Generic functions that the other scripts use. It contains some interesting functions that can be used for other purposes.

### mudify.ps1
This script Changes Blazor project to use MudBlazor, as documented at https://mudblazor.com/getting-started/installation
Just to to the directory of your .csproj and run mudify.ps1
Save your code before running this script, no guarantees!

### New-MudBlazorSolution.ps1
Creates a new MudBlazor-based Blazor Solution.

Example:
```
set-location WhateverDirectoryUnderWhichYouWantToCreateANewMudBlazorSolution;
New-MudBlazorSolution.ps1 -AppName 'Blazor Demo 03' -PageNames @('E Commerce', 'Settings') -Force;
```

## Generic Tools

### SetPriority.ps1
A simple tool that dimishes the priority of many processes, forcing them to use less computing resources and making the computer run quicker.

Processes that will be forced to the lowest priority, only being executed when the computer is idle:
- AlertusDesktop*       
- armsvc       
- Code.exe         
- EDPCleanup*           
- FileCoAuth*           
- IntuitUpdateService   
- MicrosoftSearchInBing 
- msiexec* 
- OneDrive              
- PerfWatson2           
- QualysAgent           
- RdrLeakDiag           
- Search*         
- spoolsv*              
- splwow*               
- sql*                  
- WmiPrvSE              
- TelemetryHost         
- TiWorker* 

Processes that will be forced to the below normal priority:
- CcmExec*              
- firefox*
- MSOID*                
- OfficeClickToRun      
- Nxt*                  
- policyHost            
- powershell*           
- posh       
- samsungdex*           
- SenseCE*              
- SenseNDR*  
- ServiceFabric*           
- TGitCache

Processes that many times have high priority, but this tool forces them to be executed with normal priority:
- chrome
- msedge*
- procexp*
- Teams*

