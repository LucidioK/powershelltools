# ThreatModelToVisio.ps1 -ThreatModelFilePath 'E:\dsv\ESS\internal_azure_aps\ApsThreatModel.tm7' -DiagramName 'MT / Public Cloud';
param(
    [parameter(Mandatory=$true , Position = 0)]
    [string]
    $ThreatModelFilePath,
    
    [parameter(Mandatory=$true , Position = 1)]
    [string]
    $DiagramName    
    )

function signal($n)
{
    if ($n -gt 0) { return 1; }
    if ($n -eq 0) { return 0; }
    if ($n -lt 0) { return -1; }
}

function getNodeName([object]$ell)
{
    $nel = $ell.Properties.ChildNodes | Where-Object { $_.DisplayName -eq 'Name' } | Select-Object -First 1;
    return $nel.Value."#text";
}

function GetId([object]$ell)
{
    $ell.ParentNode.Key;
}

function fixAspectRatioIfNeeded([object]$dia, [object]$vbs)
{
    function gg($elm, $tag)
    {
        $v = $elm.GetElementsByTagName($tag)[0]."#text"; 
        if ($null -eq $v)
        {
            return $null;
        }
        else 
        {
            return [float]::Parse($v);
        }
    }
    function hgt($elm) { gg $elm 'Height'; };
    function top($elm) { gg $elm 'Top';    };
    function lft($elm) { gg $elm 'Left';   };
    function wdt($elm) { gg $elm 'Width';  };

    $mxx = $dia.GetElementsByTagName('a:KeyValueOfguidanyType')            | 
            Where-Object   { $null -ne (lft $_) -and  $null -ne (wdt $_) } |
            ForEach-Object { (lft $_) + (wdt $_) }                         |
            Sort-Object                                                    |
            Select-Object -Last 1;
    $mxy = $dia.GetElementsByTagName('a:KeyValueOfguidanyType')            | 
            Where-Object   { $null -ne (hgt $_) -and  $null -ne (top $_) } |
            ForEach-Object { (hgt $_) + (top $_) }                         |
            Sort-Object                                                    |
            Select-Object -Last 1;

    $pag = $vbs.GetPage($vbs.GetPages()[0]);
    $pgs = $vbs.GetShape($pag.Name, $pag.ID);
    $pgw = $vbs.GetShapeProperty($pgs, 'PageWidth');
    $pgh = $vbs.GetShapeProperty($pgs, 'PageHeight');

    # If threat model is not in the same aspect ratio as the Visio doc, change aspect ratio of Visio doc.
    if ((signal ($mxx-$mxy)) -ne (signal ($pgw-$pgh)))
    {
        $tmp = $pgh;
        $pgh = $pgw;
        $pgw = $tmp;
        $vbs.SetProperty($pgs, 'PageWidth', $pgw);
        $vbs.SetProperty($pgs, 'PageHeight', $pgh);
    }

    $cvx = $pgw / $mxx;
    $cvy = $pgh / $mxy;
    $cvf = [Math]::Min($cvx, $cvy);

    return [PSCustomObject]@{ ConversionFactor = $cvf; PageHeight = $pgh; MaximumY =  $mxy };
}

function getFloatValue([object]$ell, [string]$tag, [float]$cvf = 1.0)
{
    $val = $ell.GetElementsByTagName($tag)[0].'#text';
    if ($null -ne $val)
    {
        return ([float]::Parse($val) * cvf);
    }

    return $null;
}

if (!(Test-Path $ThreatModelFilePath))
{
    throw "$ThreatModelFilePath not found.";
}

if ($null -eq ('VisioBase' -as [Type]))
{
    Import-Module (Join-Path $PSScriptRoot 'VisioBase.ps1');
}

[xml]$tmd = Get-Content $ThreatModelFilePath;
$dia = $tmd.GetElementsByTagName('DrawingSurfaceModel') | Where-Object { $_.Header -eq $DiagramName };

if ($null -eq $dia)
{
    throw "Could not find diagram $DiagramName on thread model $ThreatModelFilePath";
}

$vbs = [VisioBase]::NewDocument("ustrme_u.vssx");
if ($null -eq $vbs)
{
    throw "Could not start Visio. Close all Visio instances and try again.";
}

$far = fixAspectRatioIfNeeded $dia $vbs; 

$cvf = $far.ConversionFactor;
$mxy = $far.MaximumY * $cvf;
$pgh = $far.PageHeight;

$cns = $dia.GetElementsByTagName('a:Value') | Where-Object { $_.Attributes['i:type'].'#text' -eq 'Connector' } ;
$stp = @('StencilEllipse', 'StencilParallelLines', 'StencilRectangle', 'BorderBoundary', 'LineBoundary');
$els = $dia.GetElementsByTagName('a:Value') | Where-Object { $_.Attributes['i:type'].'#text' -in $stp } ;
write-host "Opened $ThreatModelFilePath, conversion factor $cvf, $($els.Count) Nodes, $($cns.Count) connectors."

$threadModelIdToShapeId = @{};
foreach ($ell in $els)
{
    $typ = $ell.Attributes['i:type'].'#text';
    $txt = getNodeName $ell;
    $hgt = getFloadValue $ell 'Height'  $cvf;       
    $wdt = getFloadValue $ell 'Width'   $cvf;
    $top = getFloadValue $ell 'Top'     $cvf;
    $lft = getFloadValue $ell 'Left'    $cvf;
    $stx = getFloadValue $ell 'SourceX' $cvf;      
    $sty = getFloadValue $ell 'SourceY' $cvf;      
    $mdx = getFloadValue $ell 'HandleX' $cvf;      
    $mdy = getFloadValue $ell 'HandleY' $cvf;      
    $enx = getFloadValue $ell 'TargetX' $cvf;      
    $eny = getFloadValue $ell 'TargetY' $cvf;      
    $top = $mxy - $top; # Converting coordinate systems: Threat Model has top left as 0,0 and Y grows downwards, while Visio has bottom left as 0,0 and Y grows upwards.
    $btm = $top - $hgt;
    $rgt = $lft + $wdt;
    switch ($typ)
    {
        'StencilEllipse' { $shp = $vbs.DrawOval($lft, $top, $rgt, $btm, $txt); }
        'LineBoundary'   { $shp = $vbs.DrawArc3Points($stx, $sty, $enx, $eny, $mdx, $mdy, $txt); }
        default          { $shp = $vbs.DrawRectangle($lft, $top, $rgt, $btm, $txt);}
    }

    if ($typ.Contains('Boundary'))
    {
        $vbs.SetProperty($shp, 'LinePattern',         '2');                  # Dashed lines
        $vbs.SetProperty($shp, 'LineColor',           $vbs.colorIndex.Red);  # Red lines
        $vbs.SetProperty($shp, 'FillPattern',         '0');                  # No fill (transparent)
        $vbs.SetProperty($shp, 'CharacterColor',       $vbs.colorIndex.Red); # Red text
    }

    if ($typ -eq 'BorderBoundary')
    {
        $vbs.SetProperty($shp, 'HorzAlign',           '2');                  # Align text right
        $vbs.SetProperty($shp, 'TxtBlkVerticalAlign', '0');                  # Align text top
    }

    $vbs.SetProperty($shp, 'CharacterFont', $vbs.Fonts['Trebuchet MS']);
    $vbs.SetProperty($shp, 'CharacterSize', '8 pt');
    $threadModelIdToShapeId[(GetId $ell)] = $shp.ID;
}

$targetToConnectorHook = @{
    East      = "Right" ;
    None      = "Bottom";
    North     = "Top"   ;
    NorthEast = "Top"   ;
    NorthWest = "Left"  ;
    South     = "Bottom";
    SouthEast = "Bottom";
    SouthWest = "Left"  ;
    West      = "Left"  ;
};

foreach ($cnt in $cns)
{
    $txt = getNodeName $cnt;
    $sid = $threadModelIdToShapeId[$cnt.SourceGuid."#text"];
    $ssh = $vbs.GetShape($null, $sid);
    $sch = $targetToConnectorHook[$cnt.PortSource."#text"];
    $tid = $threadModelIdToShapeId[$cnt.TargetGuid."#text"];
    $tsh = $vbs.GetShape($null, $tid);
    $tch = $targetToConnectorHook[$cnt.PortTarget."#text"];

    $scn = $vbs.Connect("Directed Association", $txt, $ssh, $sch, $tsh, $tch);
    $vbs.ToCurvedConnector($scn);
}

