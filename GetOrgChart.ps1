# Given a set of Active Directory user aliases, draw an Org Chart on Visio.
# - You must have Visio installed.
# - This script tries to reinstall Visio if some problem happened, but it cannot perform miracles!
# - Same with adding the ActiveDirectory module.
#
# Example:
# GetOrgChart.ps1 alias1,alias2,alias3 ...
param(
    [parameter(Mandatory=$true , Position = 0)]
    [string[]]$SeedAliases
)

class VisioBaseRectangleOnly
{
    [object]$classItem;
    [object]$interItem;
    [object]$inherItem;
    [object]$intReItem;
    [object]$compoItem;
    [Hashtable]$availableStencilItems;
    [object]$doc;
    [int]$visSelect = 2;
    [Hashtable]$colorIndex = @{
        Black   = "THEMEGUARD(RGB(0  ,  0,  0))";
        Red     = "THEMEGUARD(RGB(255,  0,  0))";
        Lime    = "THEMEGUARD(RGB(0  ,255,  0))";
        Blue    = "THEMEGUARD(RGB(0  ,  0,255))";
        Yellow  = "THEMEGUARD(RGB(255,255,  0))";
        Cyan    = "THEMEGUARD(RGB(0  ,255,255))";
        Magenta = "THEMEGUARD(RGB(255,  0,255))";
        Silver  = "THEMEGUARD(RGB(192,192,192))";
        Gray    = "THEMEGUARD(RGB(128,128,128))";
        Maroon  = "THEMEGUARD(RGB(128,  0,  0))";
        Olive   = "THEMEGUARD(RGB(128,128,  0))";
        Green   = "THEMEGUARD(RGB(0  ,128,  0))";
        Purple  = "THEMEGUARD(RGB(128,  0,128))";
        Teal    = "THEMEGUARD(RGB(0  ,128,128))";
        Navy    = "THEMEGUARD(RGB(0  ,  0,128))";
    };

    [object[]]$propertyNameAndIndices = @(
        [PSCustomObject]@{ Name = "PinX";                I = 1; J = 1;  K = 0  },
        [PSCustomObject]@{ Name = "PinY";                I = 1; J = 1;  K = 1  },
        [PSCustomObject]@{ Name = "Width";               I = 1; J = 1;  K = 2  },
        [PSCustomObject]@{ Name = "Height";              I = 1; J = 1;  K = 3  },
        [PSCustomObject]@{ Name = "LocPinX";             I = 1; J = 1;  K = 4  },
        [PSCustomObject]@{ Name = "LocPinY";             I = 1; J = 1;  K = 5  },
        [PSCustomObject]@{ Name = "Angle";               I = 1; J = 1;  K = 6  },
        [PSCustomObject]@{ Name = "FlipX";               I = 1; J = 1;  K = 7  },
        [PSCustomObject]@{ Name = "FlipY";               I = 1; J = 1;  K = 8  },
        [PSCustomObject]@{ Name = "LineWeight";          I = 1; J = 2;  K = 0  },
        [PSCustomObject]@{ Name = "LineColor";           I = 1; J = 2;  K = 1  },
        [PSCustomObject]@{ Name = "LinePattern";         I = 1; J = 2;  K = 2  },
        [PSCustomObject]@{ Name = "Rounding";            I = 1; J = 2;  K = 3  },
        [PSCustomObject]@{ Name = "EndArrowSize";        I = 1; J = 2;  K = 4  },
        [PSCustomObject]@{ Name = "BeginArrow";          I = 1; J = 2;  K = 5  },
        [PSCustomObject]@{ Name = "EndArrow";            I = 1; J = 2;  K = 6  },
        [PSCustomObject]@{ Name = "LineCap";             I = 1; J = 2;  K = 7  },
        [PSCustomObject]@{ Name = "BeginArrowSize";      I = 1; J = 2;  K = 8  },
        [PSCustomObject]@{ Name = "FillForegnd";         I = 1; J = 3;  K = 0  },
        [PSCustomObject]@{ Name = "FillBkgnd";           I = 1; J = 3;  K = 1  },
        [PSCustomObject]@{ Name = "FillPattern";         I = 1; J = 3;  K = 2  },
        [PSCustomObject]@{ Name = "ShdwForegnd";         I = 1; J = 3;  K = 3  },
        [PSCustomObject]@{ Name = "ShdwBkgnd";           I = 1; J = 3;  K = 4  },
        [PSCustomObject]@{ Name = "ShdwPattern";         I = 1; J = 3;  K = 5  },
        [PSCustomObject]@{ Name = "FillForegndTrans";    I = 1; J = 3;  K = 6  },
        [PSCustomObject]@{ Name = "FillBkgndTrans";      I = 1; J = 3;  K = 7  },
        [PSCustomObject]@{ Name = "ShdwForegndTrans";    I = 1; J = 3;  K = 8  },
        [PSCustomObject]@{ Name = "BeginX";              I = 1; J = 4;  K = 0  },
        [PSCustomObject]@{ Name = "BeginY";              I = 1; J = 4;  K = 1  },
        [PSCustomObject]@{ Name = "EndX";                I = 1; J = 4;  K = 2  },
        [PSCustomObject]@{ Name = "EndY";                I = 1; J = 4;  K = 3  },
        [PSCustomObject]@{ Name = "TheData";             I = 1; J = 5;  K = 0  },
        [PSCustomObject]@{ Name = "TheText";             I = 1; J = 5;  K = 1  },
        [PSCustomObject]@{ Name = "EventDblClick";       I = 1; J = 5;  K = 2  },
        [PSCustomObject]@{ Name = "EventXFMod";          I = 1; J = 5;  K = 3  },
        [PSCustomObject]@{ Name = "EventDrop";           I = 1; J = 5;  K = 4  },
        [PSCustomObject]@{ Name = "EventTextOverflow";   I = 1; J = 5;  K = 6  },
        [PSCustomObject]@{ Name = "LayerMember";         I = 1; J = 6;  K = 0  },
        [PSCustomObject]@{ Name = "EnableLineProps";     I = 1; J = 8;  K = 0  },
        [PSCustomObject]@{ Name = "EnableFillProps";     I = 1; J = 8;  K = 1  },
        [PSCustomObject]@{ Name = "EnableTextProps";     I = 1; J = 8;  K = 2  },
        [PSCustomObject]@{ Name = "HideForApply";        I = 1; J = 8;  K = 3  },
        [PSCustomObject]@{ Name = "PageWidth";           I = 1; J = 10; K = 0  },
        [PSCustomObject]@{ Name = "PageHeight";          I = 1; J = 10; K = 1  },
        [PSCustomObject]@{ Name = "RouteStyle";          I = 1; J = 23; K = 10; };
        [PSCustomObject]@{ Name = "ConnectorLayout";     I = 1; J = 23; K = 19 };
        [PSCustomObject]@{ Name = "TxtBlkVerticalAlign"; I = 1; J = 11; K = 4  };
        [PSCustomObject]@{ Name = "HorzAlign";           I = 4; J = 0;  K = 6  };
        [PSCustomObject]@{ Name = "CharacterColor";      I = 3; J = 0;  K = 1; };
        [PSCustomObject]@{ Name = "CharacterSize";       I = 3; J = 0;  K = 7; };
        [PSCustomObject]@{ Name = "CharacterFont";       I = 3; J = 0;  K = 0; };
    );

    [Hashtable]$propertyNameAndIndicesIndex = @{};
    [Hashtable]$Fonts = @{};

    VisioBaseRectangleOnly()
    {
        for ($i = 0; $i -lt $this.propertyNameAndIndices.Count; $i++)
        {
            $pni = $this.propertyNameAndIndices[$i];
            $this.propertyNameAndIndicesIndex.Add($pni.Name, $i);
        }
        
        if ($null -eq (Get-Command 'New-VisioDocument' -ErrorAction 'SilentlyContinue'))
        {
            Install-Module 'Visio' -Force;
            Import-Module 'Visio';
        }        
    }
    
    [VisioBaseRectangleOnly]static NewDocument([string]$stencilName = $null)
    {
        [VisioBaseRectangleOnly]::CheckIfVisioIsInstalled();
        Write-Host "Creating new drawing with stencil $stencilName" -ForegroundColor Green;
        $vb = [VisioBaseRectangleOnly]::new();
        if ([string]::IsNullOrEmpty($stencilName))
        {
            $vb.doc       = New-VisioDocument;
        }
        else 
        {
            $vb.doc       = New-VisioDocument -Stencil $stencilName;
        }
        
        $vb.Initialize();

        return $vb;
    }

    [void]static CheckIfVisioIsInstalled()
    {
        if (!([VisioBaseRectangleOnly]::IsInstalled('Visio')))
        {
            throw 'Please install Microsoft Visio.';
        }
    }

    [void] InitializePropertyNameAndIndicesIndex()
    {

        for ($i = 0; $i -lt $this.propertyNameAndIndices.Count; $i++)
        {
            $pni = $this.propertyNameAndIndices[$i];
            $this.propertyNameAndIndicesIndex.Add($pni.Name, $i);
        }
    }

    [void] Initialize()
    {
        $this.InitializeAvailableStencils();
        $this.InitializeFonts();
    }

    [void] InitializeFonts()
    {
        $this.Fonts = @{};
        foreach ($font in $this.doc.Application.ActiveDocument.Fonts)
        {
            $this.Fonts[$font.Name] = $font.ID;
        }
    }

    [void] InitializeAvailableStencils()
    {
        $this.availableStencilItems = @{};
        $docs = $this.doc.Application.Documents;

        for ($i = 1; $i -le $docs.Count; $i++)
        {
            $item = $docs.Item($i);
            if ($item.Name.EndsWith(".vssx"))
            {
                $masters = $item.Masters;
                $masters | ForEach-Object { 
                    if (!($this.availableStencilItems.ContainsKey($_.NameU))) 
                    { 
                        $this.availableStencilItems.Add($_.NameU, $_); 
                    } 
                }
            }
        }
    }

    [string[]] GetAvailableStencilItems()
    {
        $as = ($this.availableStencilItems.Keys | Select-Object | Sort-Object);
        if ($null -eq $as -or $as.Count -eq 0)
        {
            $this.InitializeAvailableStencils();
            $as = ($this.availableStencilItems.Keys | Select-Object | Sort-Object);
        }
        return $as;
    }

    [object] DrawRectangle([float]$left, [float]$top, [float]$right, [float]$bottom, [string]$title = $null)
    {
        $shape = $this.doc.Application.ActiveWindow.Page.DrawRectangle($left, $top, $right, $bottom);
        if ($null -ne $title)
        {
            $shape.Characters = $title;
        }
        return $shape;
    }

    [object] Connect([string]$connectShapeName, [string]$title, [object]$beginShape, [ConnectorHookLocation]$beginShapeHookLocation, [object]$endShape, [ConnectorHookLocation]$endnShapeHookLocation)
    {
        $this.checkIfShapeIsAvailable($connectShapeName);
        $shape     = $this.Insert($connectShapeName, $title, 0.0, 0.0, $false);
        $beginCell = $shape.Cells('BeginX');
        $endCell   = $shape.Cells('EndX');
        $beginHook = $this.getConnectHook($beginShape, $beginShapeHookLocation);
        $endHook   = $this.getConnectHook($endShape, $endnShapeHookLocation);
        $endCell.GlueTo($beginHook);
        $beginCell.GlueTo($endHook);
        return $shape;
    }

    [object] Insert([string]$shapeMasterName, [string]$title, [float]$bottom, [float]$left, [bool]$removeSubShapes = $false)
    {
        $this.checkIfShapeIsAvailable($shapeMasterName);
        $shape = $this.doc.Application.ActiveWindow.Page.Drop($this.availableStencilItems[$shapeMasterName], $bottom, $left);
        if ($null -ne $title)
        {
            $shape.Characters = $title;
        }
        return $shape;
    }

    [void] ToStraightConnector([object]$shape)
    {
        $this.SetConnectorLayout($shape, "1");
        $this.SetProperty($shape, 'RouteStyle', '16');
    }

    [void] SetConnectorLayout([object]$shape, [string]$layout)
    {
        $this.SetProperty($shape, 'ConnectorLayout', $layout);
    }

    [void] SetProperty([object]$shape, [string]$propertyName, [string]$value)
    {
        #$shape = $this.AsShape($shape);
        if ($this.propertyNameAndIndicesIndex.ContainsKey($propertyName))
        {
            $idx = $this.propertyNameAndIndicesIndex[$propertyName];
            $pni = $this.propertyNameAndIndices[$idx];   
            try
            {
                ($shape.CellsSRC($pni.I, $pni.J, $pni.K)).FormulaU = $value;     
            }
            catch
            {
                [VisioBaseRectangleOnly]::ExceptionHandler($_, $false);
            }
        }
        else
        {
            throw "Unknown property $propertyName";
        }
    }
    
    [void]static ExceptionHandler([object]$exception, [bool]$throw = $true, [string]$message = '')
    {
        [string]$cs = Get-PSCallStack;
        $cs = $cs.Replace(' at ',"`nat ");
        if ($throw)
        {
            Write-Host "`n$cs`n" -ForegroundColor Magenta;
            throw $exception.Exception.Message;
        }
        else
        {
            Write-Host "Failed:`n$($exception.Exception.Message)`n$cs" -ForegroundColor Magenta;
        }
    }

    [void] checkIfShapeIsAvailable([string]$shapeMasterName)
    {
        $as = $this.GetAvailableStencilItems();
        if (!($as.Contains($shapeMasterName)))
        {
            throw "Shape [$shapeMasterName] not available. These are the available stencil shapes: $([string]::Join(", ", $this.GetAvailableStencilItems()))";  
        }
    }

    [object] GetShape([string]$pageName, [int]$id)
    {
        try
        {
            $page = if ([string]::IsNullOrEmpty($pageName)) { $this.doc.Application.ActivePage; } else { $this.GetPage($pageName) };
            $shape = $page.Shapes.ItemFromID($id);
        }
        catch
        {
            $shape = $null;
        }
        return $shape;
    }

    [object] AsShape([object]$shape)
    {

        switch -Regex ($shape.GetType().Name)
        {  
            '__ComObject'       { return $shape; }
            'int.*|Single|long' { return $this.GetShape($null, $shape); }
            'SimplifiedShape'   { return $this.GetShape($shape.PageName, $shape.ID) }
        }
        return $null;
    }    

    [object] GetShapeProperty([object]$shape, [string]$propertyName)
    {
        #$shape = $this.AsShape($shape);
        $val = $null;
        if ($this.propertyNameAndIndicesIndex.ContainsKey($propertyName))
        {
            $idx = $this.propertyNameAndIndicesIndex[$propertyName];
            $pni = $this.propertyNameAndIndices[$idx];
            $val = ($shape.CellsSRC($pni.I, $pni.J, $pni.K)).ResultIU;
        }

        return $val;
    }

    [object] getConnectHook([object]$shape, [ConnectorHookLocation]$hookLocation)
    {
        #$shape = $this.AsShape($shape);

        if ($null -eq $shape.Master)
        {
            # If it's a base shape (circle, for instance), use the PinX cell.
            $hook = $shape.CellsSRC(1, 1, 0);
        }
        else 
        {
            $secondParam = -1;
            if ($hookLocation -eq "Top")    { $secondParam = 3; }
            if ($hookLocation -eq "Right")  { $secondParam = 0; }
            if ($hookLocation -eq "Bottom") { $secondParam = 2; }
            if ($hookLocation -eq "Left")   { $secondParam = 1; }
            $hook = $shape.CellsSRC(7, $secondParam, 0);
        }
        return $hook;
    }


    [bool]static IsInstalled([string]$appNameRegularExpression)
    {
        return ($null -ne (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object{$_.DisplayName -match $appNameRegularExpression}))
    }
}

class Person 
{
    [string]$Name;
    [string]$Alias;
    [string]$Title;
    [string]$Manager;
    [datetime]$LastJoinedIn;
    [string]$DistinguishedName;
    [Person[]]$Reports;
    [int]$Level;
    [string]$Path;
}

function simplifyTitle([string]$title)
{
    $title = $title -replace "Security ?", "Se";    
    $title = $title -replace "Site ", "Si";    
    $title = $title -replace "Rel[a-zA-Z]* ?", "Re";    
    $title = $title -replace "Software ", "Sw";
    $title = $title -replace "Eng[a-zA-Z]* ?", "En";
    $title = $title -replace "Principal ", "Pr";
    $title = $title -replace "Senior ", "Sr";
    $title = $title -replace "Partner ", "Pa";
    $title = $title -replace "Group ", "Gp";
    $title = $title -replace "Program M[a-zA-Z]* ?", "PM";
    $title = $title -replace "Manager ?", "Mn";
    $title = $title -replace "Mgr ?", "Mn";
    $title = $title -replace " ?Lead ?", "Ld";
    $title = $title -replace " ?Arch[a-zA-Z]* ?", "Ar";
    $title = $title -replace " ?Lead ?", "Ld";
    $title = $title -replace " ?Director ?", "Di";
    $title = $title.Split(",")[0];
    $title = $title.Split(":")[0];
    $title = $title.Split(" - ")[0];
    $title = $title.Trim();
    return $title;
}
function addIfNeeded([Person[]]$l, [PSCustomObject]$aduser)
{
    if ($null -eq $aduser)
    {
        Write-Host "Null user..." -ForegroundColor Yellow;
        return $l;
    }

    $item = [Person]::new();
    
    $item.Name = ($aduser.Name.Split('(')[0]).Trim();
    $item.Alias = $aduser.SamAccountName;
    $item.Title = simplifyTitle $aduser.Title;
    $item.Manager = $aduser.Manager;
    $item.LastJoinedIn = $aduser.Created;
    $item.DistinguishedName = $aduser.DistinguishedName;
    
    $existingItem = $l | Where-Object { $_.DistinguishedName -eq $item.DistinguishedName};
    if ($null -eq $existingItem)
    {
        $l += $item;
    }

    return $l;
}

function getuser([string]$id)
{
    Write-Host "Retrieving user $id" -ForegroundColor Green;
    foreach ($dc in $AllDCs)
    {
        Write-Host " Trying to get user on $dc" -ForegroundColor Green;
        $user = $null;
        try {
            $user = Get-AdUser -Identity $id -Properties Manager,Created,Title -Server $dc -ErrorAction SilentlyContinue;
        }
        catch { <# Ignore exceptions, when a user is not found in a DC Get-ADUser throws an exception, we don't care about it. #>    }
        if ($null -ne $user)
        {
            return $user;
        }
    }

    return $null;
}

function getManagementChain([Person[]]$l, [string]$id)
{
    if ([string]::IsNullOrEmpty($id))
    {
        return $l;
    }

    $checkByDistinguishedName = ($l | Foreach-Object { $_.DistinguishedName }).Contains($id);
    $checkByAlias             = ($l | Foreach-Object { $_.Alias             }).Contains($id);
    if (!$checkByDistinguishedName -and !$checkByAlias)
    {
        $aduser = getUser $id;
        $l = addIfNeeded $l $aduser;
        $l = getManagementChain $l $aduser.Manager;
    }

    return $l;
}

function addReports([Person]$person, [Person[]]$people, [int]$level)
{
    $level++;
    if (!($Script:PeoplePerLevel.ContainsKey($level)))
    { 
        $Script:PeoplePerLevel[$level] = 0;
    }

    $person.Level = $level;
    $Script:PeoplePerLevel[$level]++;
    $Script:MaxHierarchyLevel = [Math]::Max($Script:MaxHierarchyLevel, $level)
    $person.Reports = $people | Where-Object { $_.Manager -eq $person.DistinguishedName };
    foreach ($report in $person.Reports)
    {
        addReports $report $people $level;
    }
}

function insertRectangle([float]$x, [float]$y, [string]$name, [string]$alias, [string]$title)
{
    $lft = $x - ($script:RectWidth/2);
    $top = $y - ($script:RectHeight/2);
    $rgt = $lft + $script:RectWidth;
    $btm = $top + $script:RectHeight;
    $txt = "$name`n$alias`n$title";
    $rct = $script:vbs.DrawRectangle($lft, $top, $rgt, $btm, $txt);
    $script:vbs.SetProperty($rct, 'CharacterFont', $script:vbs.Fonts['Trebuchet MS']);
    $script:vbs.SetProperty($rct, 'CharacterSize', '8 pt');
    $script:vbs.SetProperty($rct, 'Rounding', '13.5 pt')
    return $rct;
}

function getByLevel([Person]$person, [int]$level, [Person[]]$l = $null)
{
    if ($null -eq $person)
    {
        return $l;
    }

    if ($null -eq $l)
    {
        $l = @();
    }

    if ($person.Level -eq $level)
    {
        $l += $person;
    }

    foreach ($report in $person.Reports)
    {
        $l = getByLevel $report $level $l;
    }

    return ($l | Sort-Object -Property Path);
}

function populatePath([Person]$person, [string]$pathUntilNow = "")
{
    if ($null -eq $person)
    {
        return;
    }

    $pathUntilNow += "/" + $person.Name;
    $person.Path = $pathUntilNow;

    foreach ($report in $person.Reports)
    {
        populatePath $report $pathUntilNow;
    }
}

function getMaxNumberOfPeopleInALevel([Person]$person )
{
    [int]$max = 0;
    for ($level = 1; $level -le $Script:MaxHierarchyLevel; $level++)
    {
        $max = [Math]::Max($max, (getByLevel $person $level).Count);
    }

    return $max;
}

function drawLineToManager([Person]$person)
{
    if (!([string]::IsNullOrEmpty($person.Manager)) -and $script:rects.ContainsKey($person.Manager))
    {
        $rectPerson  = $script:rects[$person.DistinguishedName];
        $rectManager = $script:rects[$person.Manager];
        $connector   = $vbs.Connect("Association", "", $rectPerson, "Top", $rectManager, "Bottom");
        $vbs.ToStraightConnector($connector);
    }
}

enum ConnectorHookLocation
{
    Top;
    Right;
    Bottom;
    Left;
}

function openVisio()
{
    $vbs =  [VisioBaseRectangleOnly]::NewDocument("ustrme_u.vssx");
    if ($null -eq $vbs -or $null -eq $vbs.doc -or $null -eq $vbs.doc.Application)
    {
        Write-Host "`nCould not start Visio.`nYou might need to run Office Quick Repair, which normally fixes this kind of issue.`nDo you want to run Office Quick Repair now? <Y,n>" -ForegroundColor Yellow;
        $yn = Read-Host;
        if ($yn -ne 'N')
        {
            Write-Host "`nStarting the Office Quick Repair.`nPlease follow the instruction on the Office Quick Repair dialog.`nThen retry again."  -ForegroundColor Yellow;
            $vu = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -match 'Visio'} | Select-Object -First 1).ModifyPath;
            Invoke-Expression "& $vu";
        }
        Exit-PSSession;
    }

    return $vbs;
}

function installActiveDirectoryModuleIfNeeded()
{
    IF (-not (Get-Module -Name ActiveDirectory))
    {
        if ((Get-ComputerInfo).OsProductType -eq 'Server')
        {
            Write-Host 'Module ActiveDirectory not found, installing it (Server).' -ForegroundColor Yellow;
            install-windowsfeature -Name RSAT-AD-PowerShell;
        }
        else
        {
            Write-Host 'Module ActiveDirectory not found, installing it (Workstation).' -ForegroundColor Yellow;
            $rsatPath = Join-Path $env:TEMP 'rsat.msu';
            (new-object System.Net.WebClient).DownloadFile('https://download.microsoft.com/download/1/D/8/1D8B5022-5477-4B9A-8104-6A71FF9D98AB/WindowsTH-RSAT_WS_1803-x64.msu',  $rsatPath);
            Invoke-Item -Path $rsatPath;
    
            Read-Host -Prompt 'Press enter when the installation ends.';
            Import-Module -Name ActiveDirectory -ErrorAction 'Stop' -Verbose:$false;
                
        }
    }
}

installActiveDirectoryModuleIfNeeded;

$invalidAliases = $SeedAliases | Where-Object { $_ -notmatch '^[a-z0-9]+$|^[a-z]-[a-z0-9]+$'};
if ($null -ne $invalidAliases)
{
    throw "Invalid aliases found ($([string]::Join(', ', $invalidAliases))). Aliases must follow this regular expression: '^[a-z0-9]+$|^[a-z]-[a-z0-9]+$'";
}

Write-Host "Getting DNS host... " -ForegroundColor Green -NoNewline;
$CurrentDomain = (Get-ADDomain).DnsRoot;
Write-Host $CurrentDomain -ForegroundColor Green;
Write-Host "Getting other DNS hosts... " -ForegroundColor Green -NoNewline;
$AllDCs = (Get-ADForest).Domains | 
            ForEach-Object{ (Get-ADDomainController -Filter * -Server $_).Domain } | 
            Select-Object -Unique | 
            Where-Object { $_ -ne $CurrentDomain } | 
            Sort-Object;
$AllDCs = @($CurrentDomain) + $AllDCs;
Write-Host "$($AllDCs.Count) DNS hosts found. " -ForegroundColor Green;

[Person[]]$people = @();
foreach ($alias in $SeedAliases)
{
    $aduser = getUser $alias;
    $people = addIfNeeded $people $aduser;
    $people = getManagementChain $people $aduser.Manager;
}

[float]$script:RectWidth = 1.0;
[float]$script:RectHeight = 0.5;
[float]$script:RectGap = 0.1;
[float]$script:PaperWidth = 8.5;
[float]$script:PaperHeight = 11.0;
[int]$Script:MaxHierarchyLevel = 0;
[hashtable]$Script:PeoplePerLevel = @{};
$hierarchy = $people | Where-Object { [string]::IsNullOrEmpty($_.Manager) };
addReports $hierarchy $people 0;
populatePath $hierarchy;
[int]$maxNumberOfPeopleInAnyLevel = getMaxNumberOfPeopleInALevel $hierarchy;

$script:vbs = openVisio;
$script:vbs.doc.PaperSize = 1; # Forcing "Letter" paper size.
$xMid   = ($script:PaperWidth /2);
$xWidth = $maxNumberOfPeopleInAnyLevel * ($script:RectWidth+$script:RectGap)
$xStart = $xMid - ($xWidth / 2);
$y = $script:PaperHeight;

$script:rects = @{};
for ($level = 1; $level -le $Script:MaxHierarchyLevel; $level++)
{
    $y -= ($script:RectHeight * 2);
    $peopleAtThisLevel = getByLevel $hierarchy $level;
    if ($level -eq 1 -or $peopleAtThisLevel.Count -lt 2)
    {
        $x = $xMid;
        $xInc = 0;
    }
    else 
    {
        $x    = $xStart;
        $xInc = $xWidth / ($peopleAtThisLevel.Count - 1);
    }
    
    foreach ($person in $peopleAtThisLevel)
    {
        Write-Host " Adding $($person.Name)" -ForegroundColor Green;
        $rect = insertRectangle $x $y $person.Name $person.Alias $person.Title;
        $script:rects[$person.DistinguishedName] = $rect;
        drawLineToManager $person;
        $x += $xInc;
    }
}

Write-Host "Done." -ForegroundColor Green;

return $hierarchy;