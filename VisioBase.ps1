enum ConnectorHookLocation
{
    Top;
    Right;
    Bottom;
    Left;
}

class GraphEdge
{
    [int]$ID;
    [string]$Label;
    GraphEdge(){}
    GraphEdge([int]$id){ $this.ID = $id; }
    GraphEdge([int]$id, [string]$label){ $this.ID = $id; $this.Label = $label; }
}

enum InputItemKind
{
    Lbl = 101;
    Txt = 102;
    Cmb = 103;
    Chk = 104
}

enum WindowFit
{
    None = 0;
    Page = 1;
    Width = 2;
}

class InputItem
{
    [int]$KindAsInt;
    [string]$Name;
    [string]$Label;
    [string]$Placeholder;
    [HashTable]$Options;
}

class WizardStep
{
    [string]$Name;
    [InputItem[]]$InputItems;
    [string]$OptionalConditionContextKey;
    [string]$OptionalConditionContextValue;
}

class WizardScenario
{
    [string]$Name;
    [string]$Description;
    [string]$OwnerAlias;
    [string]$IconURL;
    [string]$StyleURL;
    [WizardStep[]]$ConcreteWizardSteps;
}

class SimplifiedShape
{
    [string]$PageName;
    [int]$ID;
    [string]$ShapeMasterName;
    [string]$Text;
    [float]$Bottom;
    [float]$Left;
    [float]$Top;
    [float]$Right;
    [GraphEdge[]]$From;
    [GraphEdge[]]$To;


    SimplifiedShape(){}

    SimplifiedShape([string]$pageName, [int]$id, [string]$shapeMasterName, [string]$text)
    {
        $this.PageName = $pageName;
        $this.ID       = $id;
        $this.ShapeMasterName = $shapeMasterName;
        $this.Text     = $text;
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

    [SimplifiedShape]static FromShape([VisioBase]$vb, [object]$shape)
    {
        if ($shape -eq $null)
        {
            return $null;
        }
        else
        {
            $simplifiedShape = [SimplifiedShape]::new(
                $shape.ContainingPage.Name,
                $shape.ID,
                $shape.Master.NameU,
                ($vb.GetShapeText($shape)));

            $simplifiedShape.Bottom = [SimplifiedShape]::RoundFloat($vb.GetShapeProperty($shape, 'PinY'));
            $simplifiedShape.Left   = [SimplifiedShape]::RoundFloat($vb.GetShapeProperty($shape, 'PinX'));

            $simplifiedShape.Top    = [SimplifiedShape]::RoundFloat($simplifiedShape.Bottom + $vb.GetShapeProperty($shape, 'Height'));
            $simplifiedShape.Right  = [SimplifiedShape]::RoundFloat($simplifiedShape.Left   + $vb.GetShapeProperty($shape, 'Width' ));

            return $simplifiedShape;
        }
    }

    [SimplifiedShape]static FromShape([VisioBase]$vb, [string]$pageName, [int]$id)
    {
        $shape = $vb.GetShape($pageName, $id);
        return ([SimplifiedShape]::FromShape($vb, $shape));
    }

    [SimplifiedShape[]]static FromShapes([VisioBase]$vb, [object[]]$shapes)
    {
        $fromshapes = foreach ($shape in $shapes) { Write-Host ' ' -NoNewline -BackgroundColor Yellow; [SimplifiedShape]::FromShape($vb, $shape) }
        Write-Host;
        return $fromshapes;
    }

    [float]static RoundFloat([float]$f) { return [SimplifiedShape]::RoundFloat($f,1); }
    [float]static RoundFloat([float]$f, [int]$numberOfDigitsAfterDot) 
    { 
        [float]$md = [Math]::Pow(10, $numberOfDigitsAfterDot);
        return [float]([int]($f * $md) / $md); 
    }

    [object] ToShape([VisioBase]$vb)
    {
        return $vb.GetShape($this.PageName, $this.ID);
    }

    [void] Add([string]$toOrFrom, [int[]]$ids, [string]$label)
    {
        if ($this."$toOrFrom" -eq $null)
        {
            $this."$toOrFrom" = @();
        }
        $this."$toOrFrom" += ($ids | foreach { [GraphEdge]::new($_, $label) } );
    }
};

class VisioBase
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
        [PSCustomObject]@{ Name = "ConnectorLayout";     I = 1; J = 23; K = 19 };
        [PSCustomObject]@{ Name = "TxtBlkVerticalAlign"; I = 1; J = 11; K = 4  };
        [PSCustomObject]@{ Name = "HorzAlign";           I = 4; J = 0;  K = 6  };
        [PSCustomObject]@{ Name = "CharacterColor";      I = 3; J = 0;  K = 1; };
        [PSCustomObject]@{ Name = "CharacterSize";       I = 3; J = 0;  K = 7; };
        [PSCustomObject]@{ Name = "CharacterFont";       I = 3; J = 0;  K = 0; };
    );

    [Hashtable]$propertyNameAndIndicesIndex = @{};
    [Hashtable]$Fonts = @{};

    VisioBase()
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
    
    [VisioBase]static NewDocument([string]$stencilName = $null)
    {
        [VisioBase]::CheckIfVisioIsInstalled();
        Write-Host "Creating new drawing with stencil $stencilName" -ForegroundColor Green;
        $vb = [VisioBase]::new();
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


    [VisioBase]static OpenDocument([string]$localFilePath)
    {
        [VisioBase]::CheckIfVisioIsInstalled();
        Write-Host "Opening drawing $localFilePath" -ForegroundColor Green;
        $vb = [VisioBase]::new();
        $vb.doc       = Open-VisioDocument -Filename $localFilePath;
        $vb.Initialize();

        return $vb;
    }

    
    [VisioBase]static OpenCurrentDocument()
    {
        [VisioBase]::CheckIfVisioIsInstalled();

        Write-Host "Opening current drawing " -ForegroundColor Green;
        $vb = [VisioBase]::new();
        $vb.doc       = Get-VisioDocument -ActiveDocument;
        $vb.Initialize();

        return $vb;
    }

    [void]static CheckIfVisioIsInstalled()
    {
        if (!([VisioBase]::IsInstalled('Visio')))
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
        #$docIDs = ($docs).ID;
        #foreach ($docID in $docIDs)
        #{
        #    $item = $docs.ItemFromID($docID);
        for ($i = 1; $i -le $docs.Count; $i++)
        {
            $item = $docs.Item($i);
            if ($item.Name.EndsWith(".vssx"))
            {
                $stencilName = $item.Name;
                $masters = $item.Masters;
                $masters | foreach { 
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
        $as = ($this.availableStencilItems.Keys | select | sort);
        if ($as -eq $null -or $as.Count -eq 0)
        {
            $this.InitializeAvailableStencils();
            $as = ($this.availableStencilItems.Keys | select | sort);
        }
        return $as;
    }

    [void] FitWindow([WindowFit]$fit)
    {
        $this.doc.Application.ActiveWindow.SelectAll();
        $this.doc.Application.ActiveWindow.Selection.Move(0.1, 0.1);
        $this.doc.Application.ActiveWindow.ViewFit = $fit;
        $this.doc.Application.ActiveWindow.Selection.Move(-0.1, 0.-1);
    }

    [object] Insert([string]$shapeMasterName, [string]$title, [float]$bottom, [float]$left, [bool]$removeSubShapes = $false)
    {
        $this.checkIfShapeIsAvailable($shapeMasterName);
        $shape = $this.doc.Application.ActiveWindow.Page.Drop($this.availableStencilItems[$shapeMasterName], $bottom, $left);
        if ($null -ne $title)
        {
            $shape.Characters = $title;
        }
        if ($removeSubShapes)
        {
            $this.RemoveSubShapes($shape);
            $this.MoveTo($shape, $bottom, $left);
        }
        return $shape;
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

    [object] DrawOval([float]$left, [float]$top, [float]$right, [float]$bottom, [string]$title = $null)
    {
        $shape = $this.doc.Application.ActiveWindow.Page.DrawOval($left, $top, $right, $bottom);
        if ($null -ne $title)
        {
            $shape.Characters = $title;
        }
        return $shape;
    }

    [object] DrawArc3Points([float]$x1, [float]$y1,[float]$x2, [float]$y2,[float]$x3, [float]$y3, [string]$title = $null)
    {
        $shape = $this.doc.Application.ActiveWindow.Page.DrawArcByThreePoints($x1,$y1,$x2,$y2,$x3,$y3);
        if ($null -ne $title)
        {
            $shape.Characters = $title;
        }
        return $shape;

    }

    [object] Connect([string]$connectShapeName, [string]$title, [object]$beginShape, [ConnectorHookLocation]$beginShapeHookLocation, [object]$endShape, [ConnectorHookLocation]$endnShapeHookLocation)
    {
        $beginShape = $this.AsShape($beginShape);
        $endShape  = $this.AsShape($endShape);
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

    [void] ToStraightConnector([object]$shape)
    {
        $this.SetConnectorLayout($shape, "1");
    }

    
    [void] ToCurvedConnector([object]$shape)
    {
        $this.SetConnectorLayout($shape, "2");
    }

    [void] SetConnectorLayout([object]$shape, [string]$layout)
    {
        $this.SetProperty($shape, 'ConnectorLayout', $layout);
    }

    [void] SetLineColor([object]$shape, [string]$color)
    {
        if ($this.colorIndex.ContainsKey($color))
        {
            $color = $this.colorIndex[$color];
        }
        elseif (!($color.StartsWith('THEMEGUARD')))
        {
            throw "Unknown color $color";
        }

        $this.SetProperty($shape, 'LineColor', $color);
        foreach ($subShape in $shape.Shapes)
        {
            $this.SetProperty($subShape, 'LineColor', $color);
        }
    }

    [void] SetProperty([object]$shape, [string]$propertyName, [string]$value)
    {
        $shape = $this.AsShape($shape);
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
                [VisioBase]::ExceptionHandler($_, $false);
            }
        }
        else
        {
            throw "Unknown property $propertyName";
        }
    }

    [object] Move([object]$shape, [float]$deltaX, [float]$deltaY)
    {
        $shape = $this.AsShape($shape);
        $shape.Application.ActiveWindow.Select($shape, $this.visSelect);
        $shape.Application.ActiveWindow.Selection.Move($deltaX, $deltaY, $null);
        return $shape;
    }

    
    [object] MoveTo([object]$shape, [float]$x, [float]$y)
    {
        $shape = $this.AsShape($shape);
        $currentX = $this.GetShapeProperty($shape, 'PinX');
        $currentY = $this.GetShapeProperty($shape, 'PinY');
        $deltaX   = $x - $currentX;
        $deltaY   = $y - $currentY;
        $this.Move($shape, $deltaX, $deltaY);
        return $shape;
    }

    [object] RemoveSubShapes([object]$shape)
    {
        $shape = $this.AsShape($shape);
        $this.GetSubShapes($shape) | foreach { 

            $subShape = $shape.ContainingPage.Shapes.ItemFromID($_);
            $subShape.RemoveFromContainers();
            $subShape.Delete() 
        };

        return $shape;
    }

    [object[]] GetSubShapes([object]$shape)
    {
        $shape = $this.AsShape($shape);
        $ss = $null;
        try
        {
            $ss = $shape.ContainerProperties.GetListMembers();
        }
        catch
        {
            $id = $shape.ID;
            $ss = $this.GetShapes($shape.ContainingPage.NameU, $null) | where { $_.MemberOfContainers.Contains($id) };
        }
        return $ss;
    }


    [SimplifiedShape[]] GetSubSimplifiedShapes([object]$shape)
    {
        Write-Host "Getting simplified shapes, this might take some time... " -ForegroundColor Green -NoNewline;
        $shape = $this.AsShape($shape);
        $ss = $null;
        try
        {
            $ss = $shape.ContainerProperties.GetListMembers();
        }
        catch
        {
            $id = $shape.ID;
            $ss = $this.GetShapes($shape.ContainingPage.NameU, $null) | where { $_.MemberOfContainers.Contains($id) };
        }
        $ss = [SimplifiedShape]::FromShapes($this, $ss) | sort  -Property @{ expression = "Bottom"; Descending = $True },@{ expression = "Left"; Descending = $False };
        return $ss;
    }

    [void] Delete([object]$shape)
    {
       $shape = $this.AsShape($shape);
       $shape.Delete();
    }

    [void] SaveAs([string]$filePath)
    {
        Save-VisioDocument -Filename $filePath -Document $this.doc;
    }


    [void] Close()
    {
        $this.doc.Close();
        Close-VisioApplication;
    }

    [void] checkIfShapeIsAvailable([string]$shapeMasterName)
    {
        $as = $this.GetAvailableStencilItems();
        if (!($as.Contains($shapeMasterName)))
        {
            throw "Shape [$shapeMasterName] not available. These are the available stencil shapes: $([string]::Join(", ", $this.GetAvailableStencilItems()))";  
        }
    }

    [object] GetShapeProperty([object]$shape, [string]$propertyName)
    {
        $shape = $this.AsShape($shape);
        $val = $null;
        if ($this.propertyNameAndIndicesIndex.ContainsKey($propertyName))
        {
            $idx = $this.propertyNameAndIndicesIndex[$propertyName];
            $pni = $this.propertyNameAndIndices[$idx];
            $val = ($shape.CellsSRC($pni.I, $pni.J, $pni.K)).ResultIU;
        }

        return $val;
    }

    [PsCustomObject[]] GetAllShapeProperties([object]$shape)
    {
        $shape = $this.AsShape($shape);
        $props = @();
        for ($i = 1; $i -lt 16; $i++)
        {
            for ($j = 1; $j -lt 16; $j++)
            {
                for ($k = 0; $k -lt 16; $k++)
                {
                    if ($shape.CellsSRCExists(0, $i, $j, $k) -ge 0)
                    {
                        $p = $null;
                        try
                        {
                            $p = $shape.CellsSRC($i, $j, $k);
                        }
                        catch
                        {
                            Write-Host "Failed with $($_.Exception.Message) $i $j $k" -ForegroundColor Magenta;
                            continue;
                        }
                        if ($p -ne $null -and $p.Name -match '^[A-Za-z]+$' -or $p.Name.StartsWith('Connections.'))
                        {
                            $props += [PsCustomObject]@{ Name =  $p.Name; Value = $p.ResultIU; I=$i; J = $j; K = $k };  
                        }
                        else
                        {
                            break;
                        }
                    }
                    else
                    {
                        break;
                    }
                }
            }
        }
        return $props;
    }

    [object] GetDocumentItem([int]$index)
    {
        try
        {
            $itemFromID = $this.doc.Application.Documents.ItemFromID($index);
        }
        catch
        {
            $itemFromID = $null;
        }

        return $itemFromID;
    }

    [object[]] GetPages()
    {
        $pages = $this.doc.Pages | foreach { $_.Name };
        return $pages;
    }

    [object] GetPage([string]$name)
    {
        $page = $this.doc.Pages | where { $_.Name -eq $name};
        return $page;
    }

    [object[]] GetShapes([string]$pageName, [string]$shapeMasterName)
    {
        $shapes = @();
        $pageList = if ([string]::IsNullOrEmpty($pageName)) { $this.GetPages() } else { @( ($this.GetPage($pageName)) ) };
        foreach ($page in $pageList)
        {
            if (($page.GetType()).Name -eq 'string')
            {
                $pageShapes = ($this.doc.Pages.Item($page)).Shapes;
            }
            else
            {
                $pageShapes = $page.Shapes;
            }

            if (!([string]::IsNullOrEmpty($shapeMasterName)))
            {
                $pageShapes = $pageShapes | where { $_.Master.NameU -eq $shapeMasterName };
            }

            $shapes += $pageShapes;
        }
        
        return $shapes;
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

    [string] GetShapeText([object]$shape)
    {
        $shape = $this.AsShape($shape);
        $type = $shape.GetType();
        if ($type.Name -eq 'PSCustomObject' -and $shape.Text -ne $null)
        {
            return $shape.Text;
        }
        else
        {
            $shapeText = '';
            if ([string]::IsNullOrEmpty($shape.Characters.TextAsString))
            {
                $textList  = $shape.Shapes | foreach { $_.Characters.TextAsString };
                if ($textList -ne $null)
                {
                    $shapeText = [string]::Join("`n" ,$textList);
                }
            }
            else
            {
                $shapeText = $shape.Characters.TextAsString;
            }
            $shapeText = ($shapeText | ConvertTo-Json).Replace('\u2028', "`n") | convertfrom-json;
            return $shapeText;
        }
    }

    [SimplifiedShape] GetSimplifiedShape([string]$pageName, [int]$id)
    {
        return ([SimplifiedShape]::FromShape($this, $pageName, $id));
    }


    [SimplifiedShape] GetSimplifiedShape([object]$shape)
    {
        if ($shape.GetType().Name -eq 'SimplifiedShape')
        {
            return $shape;
        }

        $shape = $this.AsShape($shape);
        return ([SimplifiedShape]::FromShape($this, $shape));
    }

    [SimplifiedShape] GetRelatedShapes([object]$shape)
    {
        $shape = $this.AsShape($shape);
        $fromOrTos = @('From', 'To');
        [SimplifiedShape]$r = $this.GetSimplifiedShape($shape);

        foreach ($fromOrTo in $fromOrTos)
        {
            $opposite = if ($fromOrTo -eq 'From') { 'To' } else { 'From' };
            $id       = $shape.ID;
            $connects = $shape."$($fromOrTo)Connects";
            foreach ($connect in $connects)
            {
                $connectedShape = $connect."$($fromOrTo)Cell".Shape;
                if ($connectedShape.Master.NameU -ne 'Dynamic connector')
                {
                    throw "I only know how to deal with relationships through Dynamic Connector, this is using $($connectedShape.Master.NameU)";
                }
                $connectorText = $this.GetShapeText($connectedShape);
                $connectorId   = $connectedShape.Id;

                $hasEndArrow   = $this.GetShapeProperty($connectedShape, 'EndArrow') -ne 0;
                $pointedShapeIndex = if ($hasEndArrow) {2} else {1};
                $relatedId     = $connectedShape.Connects.Item($pointedShapeIndex)."$($opposite)Cell".Shape.ID;
                if ($relatedId -ne $id)
                {
                    $r.Add("$opposite", $relatedId, $connectorText);
                }
            }
        }
        return $r;
    }

    [object[]] GetGraphStartingAt([object]$shape)
    {
        $shape = $this.AsShape($shape);
        return $this.GetGraphStartingAt($shape.ContainingPage.NameU, $shape);  
    }

    [object[]] GetGraphStartingAtInitialState()
    {
        $is = $this.GetShapes($null, 'Initial state');
        if ($is -eq $null -or $is.Count -eq 0)
        {
            throw "No Initial State shape found.";
        }

        if ($is.Count -gt 1)
        {
            throw "More than one Initial State shape found. Please use the version of GetGraphStartingAtInitialState where you can determine which page to inspect, or delete unused Initial State shapes.";
        }

        $is = $is[0];
        return ($this.GetGraphStartingAt($is.ContainingPage.Name, $is));
    }

    
    [object[]] GetGraphStartingAtInitialState([string]$pageName)
    {
        $is = $this.GetShapes($pageName, 'Initial state');
        if ($is -eq $null -or $is.Count -eq 0)
        {
            throw "No Initial State shape found in page $pageName.";
        }

        if ($is.Count -gt 1)
        {
            throw "More than one Initial State shape found in page $pageName.";
        }

        $is = $is[0];
        return ($this.GetGraphStartingAt($pageName, $is));
    }

    [object[]] GetGraphStartingAt([string]$pageName, [object]$shape)
    {
        $shape = $this.AsShape($shape);
        Write-Host 'GetGraphStartingAt - starting' -ForegroundColor Green;
        if ($shape.GetType().Name.StartsWith('Int'))
        {
            $shape = $this.GetShape($pageName, $shape);
        }

        $alreadyVisitedIds = [System.Collections.Generic.HashSet[int]]::new();

        $currentLevel = @( $shape );
        $graph = @( );
        while ($currentLevel.Count -gt 0)
        {
            $nextLevel = @();
            foreach ($currentLevelShape in $currentLevel)
            {
                if (($graph | where { $_.ID -eq $currentLevelShape.ID }) -ne $null)
                {
                    continue;
                }

                $alreadyVisitedIds += $currentLevelShape.ID;
                [SimplifiedShape]$relatedShapes = $this.GetRelatedShapes($currentLevelShape);
                Write-Host " [$($relatedShapes.ShapeMasterName)] [$($relatedShapes.Text.Split("`n")[0])]" -ForegroundColor Green;
                $graph += $relatedShapes;
                $fromOrTos = @('From', 'To');
                foreach ($fromOrTo in $fromOrTos)
                {
                    foreach ($edge in $relatedShapes."$fromOrTo")
                    {
                        $alreadyVisitedIds += $edge.ID;
                        $nextLevel         += $this.GetShape($relatedShapes.PageName, $edge.ID);
                    }
                }
            }

            $currentLevel = $nextLevel;
        }
        Write-Host "GetGraphStartingAt - done, $($graph.Count) nodes on graph" -ForegroundColor Green;
        return $graph;
    }

    [PSCustomObject] GetProperties([object]$shape)
    {
        $shape = $this.AsShape($shape);
        $ret = @{};
        foreach ($pni in $this.propertyNameAndIndices)
        {
            $p = $this.GetShapeProperty($shape, $pni.Name);#$shape.CellsSRC($pni.I, $pni.J, $pni.K);
            $ret.Add($pni.Name, $p.ResultIU);
        }
        $pco = [PSCustomObject]$ret;
        return $pco;
    }

    [object] getConnectHook([object]$shape, [ConnectorHookLocation]$hookLocation)
    {
        $shape = $this.AsShape($shape);

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

    [WizardScenario] WizardScenarioFromVisioDiagram([string]$WizardScenarioName)
    {
        return ($this.WizardScenarioFromVisioDiagram($WizardScenarioName, $null, $null, $null, $null, $null));
    }

    [WizardScenario] WizardScenarioFromVisioDiagram(
        $WizardScenarioName       ,
        $WizardScenarioDescription = $null,
        $WizardScenarioOwnerAlias  = $null,
        $WizardScenarioIconURL     = $null,
        $WizardScenarioStyleURL    = $null,
        $PageName                  = $null
        )

    {
        function DescriptionToName
        {
            [OutputType([string])]
            param([string]$d)

            $words = (($d -replace '[^A-Za-z 0-9]',' ') -replace ' +',' ').Split(' ') | 
                where { $_.Length -gt 0 } | 
                foreach { [char]::ToUpperInvariant($_[0]) + $_.SubString(1) };

            return [string]::Join("", $words);
        }

        function IsLbl
        {
            [OutputType([bool])]
            param([SimplifiedShape]$ss)

            return ([string]::IsNullOrEmpty($ss.ShapeMasterName) -or $ss.ShapeMasterName -eq 'Label');
        }

        function IsTxt
        {
            [OutputType([bool])]
            param([SimplifiedShape]$ss)

            return ($ss.ShapeMasterName -eq "Text box");
        }

        function IsCmb
        {
            [OutputType([bool])]
            param([SimplifiedShape]$ss)

            return ($ss.ShapeMasterName -eq 'Combo Box' -or $ss.ShapeMasterName -eq 'Drop down');
        }

        function IsChk
        {
            [OutputType([bool])]
            param([SimplifiedShape]$ss)

            return ($ss.ShapeMasterName -eq 'Checkbox');
        }

        function IsBtn
        {
            [OutputType([bool])]
            param([SimplifiedShape]$ss)

            return ($ss.ShapeMasterName -eq 'Button');
        }

        function GetDialogDefinition
        {
            [OutputType([WizardStep])]
            param([SimplifiedShape[]]$graph, [int]$graphIndex)
    
            function getCmb
            {
                [OutputType([InputItem])]
                param([SimplifiedShape[]]$ss, [InputItem]$ii)

                $ii.KindAsInt = [int]([InputItemKind]::Cmb);
                $optionsList = $null;
                if (!([string]::IsNullOrEmpty($ss.Text)))
                {
                    $optionsList = $ss.Text.Split("`n") | where { !([string]::IsNullOrEmpty($_)) } | select -Unique;
                }

                if ($optionsList -eq $null)
                {
                    $optionsList = @("No", "Options", "Informed");
                }

                $ii.Placeholder = $optionsList[0];
                $options = @{};
                $optionsList | foreach { $options.Add($_, $_); }
                $ii.Options = $options;
                return $ii;
            }

            function getInputItem
            {
                [OutputType([InputItem])]
                param([SimplifiedShape[]]$ss)

                if ([string]::IsNullOrEmpty($ss[0].Text))
                {
                    return $null;
                }

                [InputItem]$ii = [InputItem]::new();
                $ii.Label = $ss[0].Text;
                $ii.Name  = DescriptionToName $ss[0].Text;
                if ($ss.Count -eq 1 -and (IsLbl $ss[0]))
                {
                    $ii.KindAsInt = [int]([InputItemKind]::Lbl);
                }
                elseif ($ss.Count -eq 2 -and (IsLbl $ss[0]) -and (IsTxt $ss[1]))
                {
                    $ii.KindAsInt = [int]([InputItemKind]::Txt);
                    $ii.Placeholder = $ss[1].Text;
                }
                elseif ($ss.Count -eq 2 -and (IsLbl $ss[0]) -and (IsCmb $ss[1]))
                {
                    $ii = getCmb $ss[1] $ii;
                }
                elseif ($ss.Count -eq 2 -and (IsLbl $ss[0]) -and (IsChk $ss[1]))
                {
                    $ii.KindAsInt = [int]([InputItemKind]::Chk);
                }
                elseif ($ss.Count -eq 3 -and (IsTxt $ss[1]) -and (IsCmb $ss[2]))
                {
                    $ii = getCmb $ss[1] $ii;
                }
                elseif ($ss.Count -eq 2 -and (IsBtn $ss[0]) -and (IsBtn $ss[1]) -and $ss[0].Text -eq 'Previous' -and $ss[1].Text -eq 'Next')
                {
                    return $null;
                }
                elseif ($ss.Count -eq 1 -and ($ss[0].Text -eq 'Previous' -or $ss[1].Text -eq 'Next'))
                {
                    return $null;
                }
                else
                {
                    Write-Host "`nFound weird combination of $($ss.Count): $([string]::Join(',', ($ss).ShapeMasterName))`n" -ForegroundColor Magenta;
                    return $null;
                }

                return $ii;
            }

            function isIntersecting
            {
                [OutputType([bool])]
                param([SimplifiedShape]$s1, [SimplifiedShape]$s2)

                $m2 = (($s2.Top + $s2.Bottom) / 2);
                $t1 = $s1.Top;
                $b1 = $s1.Bottom;
                $isIt = ($s1.Top -gt $m2 -and $s1.Bottom -lt $m2);

                return $isIt;
            }

            function getIntersectingShapes
            {
                [OutputType([SimplifiedShape[]])]
                param([SimplifiedShape[]]$ss, [int]$pos)

                $is = @();
                $is += $ss[$pos];
                for ($i = $pos + 1; $i -lt $ss.Count; $i++)
                {
                    if (isIntersecting $ss[$pos] $ss[$i])
                    {
                        $is += $ss[$i];
                    }
                    else
                    {
                        break;
                    }
                }
                return $is;
            }

            function getPredecessorIfExists([SimplifiedShape[]]$graph, [int]$graphIndex)
            {
                $currentNodeId = $graph[$graphIndex].ID;
                for ($i = $graphIndex -1; $i -ge 0; $i--)
                {
                    $referencesToCurrentNodeId = $graph[$i].To | where { $_.ID -eq $currentNodeId };
                    if ($referencesToCurrentNodeId -ne $null -and $referencesToCurrentNodeId.Count -gt 0)
                    {
                        return $graph[$i];
                    }
                }
                return $null;
            }

            
            function setDialogConditionIfExists([WizardStep]$wizardStep, [SimplifiedShape[]]$graph, [int]$graphIndex)
            {
                [SimplifiedShape]$predecessor = getPredecessorIfExists $graph $graphIndex;
                if ($predecessor -ne $null)
                {
                    [string]$edgeLabel = ($predecessor.To | where { $_.ID -eq $graph[$graphIndex].ID } | select -First 1).Label;
                    if (!([string]::IsNullOrEmpty($edgeLabel)) -and $edgeLabel.Contains(':'))
                    {
                        $split = $edgeLabel.Split(':');
                        $wizardStep.OptionalConditionContextKey = $split[0];
                        $wizardStep.OptionalConditionContextValue = $split[1];
                    }
                }

                return $wizardStep;
            }

            $containerShape = $graph[$graphIndex];
            [SimplifiedShape[]]$subShapes = $this.GetSubSimplifiedShapes($containerShape);

            if ($subShapes -eq $null -or $subShapes.Count -eq 0 -or !(IsLbl $subShapes[0]))
            {
                write-host "This step does not have a title, neeeext!" -ForegroundColor Magenta;
                return $null;
            }
            $title     = $subShapes[0].Text;
            $wizardStep = [WizardStep]::new();
            Write-Host "Building step dialog [$title]" -ForegroundColor Green;
            $wizardStep.Name = DescriptionToName $title;
            $wizardStep = setDialogConditionIfExists $wizardStep $graph $graphIndex;
            $wizardStep.InputItems = @();
            for ($i = 0; $i -lt $subShapes.Count; $i++)
            {
                $iss = getIntersectingShapes $subShapes $i;
                $ii  = getInputItem $iss;
                if ($ii -ne $null)
                {
                    $wizardStep.InputItems += $ii;
                }

                $i += $iss.Count - 1;
            }

            return $wizardStep;
        }

        Write-Host "`n`nBuilding Wizard Scenario from Visio Diagram" -ForegroundColor Green;

        [SimplifiedShape[]]$graph = $this.GetGraphStartingAtInitialState($PageName);
        [WizardScenario]$wizardScenario = [WizardScenario]::new();
        $wizardScenario.Name        = $WizardScenarioName       ;
        $wizardScenario.Description = $WizardScenarioDescription;
        $wizardScenario.OwnerAlias  = $WizardScenarioOwnerAlias ;
        $wizardScenario.IconURL     = $WizardScenarioIconURL    ;
        $wizardScenario.StyleURL    = $WizardScenarioStyleURL   ;
        $wizardScenario.ConcreteWizardSteps = @();

        for ($i = 1; $i -lt $graph.Count; $i++)
        {
            $gd = GetDialogDefinition $graph $i;
            if ($gd -ne $null)
            {
                $wizardScenario.ConcreteWizardSteps += $gd;
            }
        }

        return $wizardScenario;

    }


    [void]static Test()
    {
        del "$(join-path $env:LOCALAPPDATA 'Microsoft\Visio')\*.vs*";
        Get-Process 'Visio' -ErrorAction SilentlyContinue | foreach { $_.Kill() };
        $global:vb = [VisioBase]::NewDocument("ustrme_u.vssx");
        $global:c1 = $global:vb.Insert('Class', 'BaseClass', 2, 8, $true);
        $global:ps = $global:vb.GetProperties($global:c1);
        $x = $global:vb.GetShapeProperty($global:c1, "PinX");
        $y = $global:vb.GetShapeProperty($global:c1, "PinY");
        $w = $global:vb.GetShapeProperty($global:c1, "Width");
        $h = $global:vb.GetShapeProperty($global:c1, "Height");
        Write-Host "x: $x y: $y w: $w h: $h";

        $global:c2 = $global:vb.Insert('Class', 'Inherited', 2, 4, $true);
        $global:in = $global:vb.Connect('Inheritance', $null, $global:c1, "Bottom", $global:c2, "Top");
        $fn = Join-Path "$($env:USERPROFILE)\documents" "test.vsdx";
        Remove-Item -Path $fn -ErrorAction SilentlyContinue;
        $global:vb.SaveAs($fn);
        $global:vb.Close();

        $global:vb = [VisioBase]::OpenDocument($fn);
        $global:pages = $global:vb.GetPages() | convertto-json -Depth 1 | ConvertFrom-Json;
        $global:allShapes = $global:vb.GetShapes($null, $null) | convertto-json -Depth 1 | ConvertFrom-Json;
        $global:classShapes = $global:vb.GetShapes($null, 'Class');
        $global:inheritanceShapes = $global:vb.GetShapes($null, 'Inheritance');
    }

    [void]static TestWizardScenarioCreation([string]$visioDiagramPath, [string]$JsonOutputPath)
    {
        del "$(join-path $env:LOCALAPPDATA 'Microsoft\Visio')\*.vs*";
        Get-Process 'Visio' -ErrorAction SilentlyContinue | foreach { $_.Kill() };
        $global:vb = [VisioBase]::OpenDocument($visioDiagramPath);
        $global:WizardScenario = $global:vb.WizardScenarioFromVisioDiagram('Some Test');
        $global:WizardScenario | ConvertTo-Json -Depth 12 -Compress | out-file $JsonOutputPath -Encoding ascii;
    }

    [void]static TestWizardScenarioCreation()
    {
        $vsdxPath = Join-Path $PSScriptRoot 'dialogs.vsdx';
        $jsonPath = Join-Path $PSScriptRoot 'WizardScenario.json';
        [VisioBase]::TestWizardScenarioCreation($vsdxPath, $jsonPath);
    }

    [void]static TestWizardScenarioCreationFromCurrentDocument([string]$JsonOutputPath)
    {
        $global:vb = [VisioBase]::OpenCurrentDocument();
        $global:WizardScenario = $global:vb.WizardScenarioFromVisioDiagram('Some Test');
        $global:WizardScenario | ConvertTo-Json -Depth 12 -Compress | out-file $JsonOutputPath -Encoding ascii;        
    }

    
    [void]static TestWizardScenarioCreationFromCurrentDocument()
    {
        [VisioBase]::TestWizardScenarioCreationFromCurrentDocument((Join-Path $PSScriptRoot 'WizardScenario.json'));
    }

    [bool]static IsInstalled([string]$appNameRegularExpression)
    {
        return ((Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | ?{$_.DisplayName -match $appNameRegularExpression}) -ne $null)
    }
    
    [string]static GetVisioAppLocation()
    {
        $il = Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall  | % { Get-ItemProperty $_.PsPath }  | where { $_.DisplayName -ne $null } | where { $_.DisplayName.Contains('Microsoft Visio') } | Select -First 1 InstallLocation | select;
        $rt = Join-Path $il.InstallLocation 'root';
        $op = (get-childitem -Path $rt -Filter 'Office*').FullName | sort | select -Last 1;
        $vl = Join-Path $op 'Visio.exe';
        if (!(Test-Path $vl))
        {
            throw "Visio not found at $vl";
        }

        return $vl;
    }

    [string]static ReadFileFromZip([string]$archivePath, [string]$fileName)
    {
        $fileZip = [System.IO.Compression.ZipFile]::Open($archivePath, 'Update');
        $entry   = $fileZip.Entries | Where-Object { $_.FullName -match $fileName } | select -First 1 | select;
        $reader  = [System.IO.StreamReader]($entry).Open();
        $content = $reader.ReadToEnd();
        $reader.Close();
        $reader.Dispose();
        $fileZip.Dispose();
        return $content;
    }

    [string[]]static GetXMLItemsTextFromItemInZip([string]$archivePath, [string]$fileName, [string]$xmlTag)
    {
        [xml]$xml = [VisioBase]::ReadFileFromZip($archivePath, $fileName);
        $txt = ($xml.GetElementsByTagName($xmlTag)).'#text';
        return $txt;
    }

    [object[]]static GetVisioAvailableStencils()
    {
        $vl = [VisioBase]::GetVisioAppLocation();
        $vf = [System.IO.Path]::GetDirectoryName($vl);
        $vc = Join-Path $vf "Visio Content";
        $vc = Join-Path $vc ((Get-Culture).LCID);
        $fs = (Get-ChildItem -Path $vc -Filter "*.vssx").FullName;
        $rs = @();
        $tt = $fs.Count;
        $ct = 0.0;
        foreach ($f in $fs)
        {
            $ct++;
            $nm = ([System.IO.Path]::GetFileName($f));
            Write-Progress -PercentComplete ($ct * 100 / $tt) -Activity $nm;
            $dsc = [VisioBase]::GetXMLItemsTextFromItemInZip($f, 'core.xml', 'dc:title');
            $masterShapes = [VisioBase]::GetXMLItemsTextFromItemInZip($f, 'app.xml', 'vt:lpstr');
            $rs  += [PSCustomObject]@{ 
                Description = $dsc; 
                Name = $nm; 
                FullPath = $f;
                MasterShapes = $masterShapes
            };
        }
        return $rs;
    }
}