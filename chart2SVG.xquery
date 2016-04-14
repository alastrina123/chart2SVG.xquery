xquery version "3.0";

declare namespace xs="http://www.w3.org/2001/XMLSchema";
declare namespace c="http://schemas.openxmlformats.org/drawingml/2006/chart";
declare namespace a="http://schemas.openxmlformats.org/drawingml/2006/main";
declare namespace b="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
declare namespace r="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
declare namespace cdr="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing";
declare namespace xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
declare namespace rels="http://schemas.openxmlformats.org/package/2006/relationships";
declare namespace math="http://www.w3.org/2005/xpath-functions/math";
declare default element namespace "http://www.w3.org/2000/svg";

declare boundary-space strip;

declare namespace output = "http://www.w3.org/2010/xslt-xquery-serialization"; 
declare option output:method "xml"; 
declare option output:indent "yes"; 
declare option output:omit-xml-declaration "no";
declare option output:encoding "UTF-8";

declare variable $ptCMfactor := 0.035277777777777776;

declare variable $stdFONT as xs:string external := 'Arial';
declare variable $stdFONTptSZ as xs:decimal external;
declare variable $stdFONTcol as xs:string external :='black';

declare variable $axlabFONTcol as xs:string external :='black';

declare variable $pathSTRK as xs:float external;
declare variable $pathCOL as xs:string external;

declare variable $gridCOL as xs:string external := 'rgb(159,159,159)';

declare variable $sheetNAME as xs:string external := 'DATASHEET CURVERS';

declare variable $sheetrID := //b:sheet[@name=$sheetNAME]/@r:id/data();
declare variable $sheetTGT := doc(concat(substring-before(document-uri(.),'workbook.xml'),'_rels/workbook.xml.rels'))//rels:Relationship[@Id=$sheetrID]/@Target/data();
declare variable $sheetFILE := doc(concat(substring-before(document-uri(.),'workbook.xml'),$sheetTGT));

declare variable $sheetDWGSrID := $sheetFILE//b:drawing/@r:id/data();
declare variable $sheetRELSFILE := doc(concat(substring-before(document-uri(.),'workbook.xml'),'worksheets/_rels',substring-after($sheetTGT,'worksheets'),'.rels'));
declare variable $sheetDWGSTGT := $sheetRELSFILE//rels:Relationship[@Id=$sheetDWGSrID]/@Target/data();
declare variable $sheetDWGSFILE := doc(concat(substring-before(document-uri(.),'workbook.xml'),substring-after($sheetDWGSTGT,'../')));

declare variable $colWIDTHdef external := 9.14;
declare variable $rowHEIGHTdef external := 12.75;

declare variable $chartAREAwd external;
declare variable $chartAREAht external;

declare variable $plotAREAwd external;
declare variable $plotAREAht external;

declare function local:svgTITLE($TITLE)
    {element title
        {text {$TITLE}
    }
};

declare function local:svgPATH($pathID,$pathCOL, $pathSTRK, $dXcoord, $dYcoord)
    {element path
        {attribute id {$pathID},
         attribute stroke {$pathCOL},
         attribute stroke-width {$pathSTRK},
         attribute fill {'none'},
         attribute d {
            if (contains($pathID,'smooth'))
            then let $count := 0
            for sliding window $xW in ($dXcoord)
                start at $pt0 when fn:true()
                only end at $pt2 when $pt2 - $pt0 eq 2

                let $x0 := $xW[1]
                let $x1 := $xW[2]
                let $x2 := $xW[3]
                let $y0 := $dYcoord[$pt0]
                let $y1 := $dYcoord[$pt0 + 1]
                let $y2 := $dYcoord[$pt0 + 2]         
                let $tFACT := 1 div 3
                let $Dp0p1 := math:pow(math:pow($x1 - $x0, 2) + math:pow($y1 - $y0, 2), 0.5)
                let $Dp1p2 := math:sqrt(math:pow($x2 - $x1,2) + math:pow($y2 - $y1,2))
                let $C1x := $x1 - (($Dp0p1 div ($Dp0p1 + $Dp1p2)) * ($x2 - $x0) * $tFACT)
                let $C1y := $y1 - (($Dp0p1 div ($Dp0p1 + $Dp1p2)) * ($y2 - $y0) * $tFACT)
                let $C2x := $x1 + (($Dp1p2 div ($Dp0p1 + $Dp1p2)) * ($x2 - $x0) * $tFACT)
                let $C2y := $y1 + (($Dp1p2 div ($Dp0p1 + $Dp1p2)) * ($y2 - $y0) * $tFACT)
                return if ($pt0 eq 1) then concat('M',$x0,' ',$y0,'Q') else concat(' ',$C1x,' ',$C1y,' ',$x1,' ',$y1,' ', if ($pt2 eq count($dYcoord)) then concat('Q',$C2x,' ',$C2y,' ',$x2,' ',$y2) else concat('C',$C2x,' ',$C2y))
            else for $pts at $idx in $dXcoord
            return concat(if ($idx=1) then 'M' else 'L',' ',$dXcoord[$idx],' ',$dYcoord[$idx])}
    }
};

declare function local:svgRECT($rectXorg,$rectYorg,$rectWD,$rectHT,$rectFILLcol,$rectSTKcol,$rectSTKwd)
    {element rect
        {attribute x {$rectXorg},
        attribute y {$rectYorg},
        attribute width {$rectWD},
        attribute height{$rectHT},
        attribute fill{$rectFILLcol},
        attribute stroke{$rectSTKcol},
        attribute stroke-width{$rectSTKwd}
    }
};

declare function local:svgTEXT($textID,$txtSTYLE, $txtFONT, $txtSIZE, $txtCOL,$txtXpos,$txtYpos,$txtLINE)
    {element text
        {attribute id {$textID},
        attribute x {$txtXpos},
        attribute y {$txtYpos},
        attribute style {$txtSTYLE},
        attribute font-family {$txtFONT},
        attribute font-size {$txtSIZE * $ptCMfactor},
        attribute fill {$txtCOL},
        if (contains($textID,'txtbox') or contains($textID,'axtitle')) then
            for $txtRUN in $txtLINE/a:r
            return
        element tspan
        {attribute baseline-shift
        {switch ($txtRUN/a:rPr/@baseline/data())
            case '-25000' return ('sub')
            case '25000' return ('sup')
            default return ('')},
        attribute font-size
        {switch ($txtRUN/a:rPr/@baseline/data())
            case () return ($txtSIZE * $ptCMfactor)
            default return ($txtSIZE * $ptCMfactor * 0.8)},
        text {$txtRUN}}
        else text {$txtLINE}
    }
};

declare function local:chartWIDTH()
    {for $chartNAMES in $sheetDWGSFILE//xdr:cNvPr[contains(@name, 'Chart')]
    let $chartFRcol as xs:integer := $chartNAMES/ancestor::xdr:graphicFrame/preceding-sibling::xdr:from/xdr:col/text() cast as xs:integer
    let $chartFOFFcol as xs:integer := $chartNAMES/ancestor::xdr:graphicFrame/preceding-sibling::xdr:from/xdr:coloff/text() cast as xs:integer

    let $chartTOcol as xs:integer := $chartNAMES/ancestor::xdr:graphicFrame/preceding-sibling::xdr:to/xdr:col/text() cast as xs:integer
    let $chartTOFFcol as xs:integer := $chartNAMES/ancestor::xdr:graphicFrame/preceding-sibling::xdr:to/xdr:coloff/text() cast as xs:integer

    for $chartCOLUMNS in $chartFRcol to $chartTOcol
        for $colWIDTHS in $sheetFILE//col
        where ($chartCOLUMNS ge $colWIDTHS//col/@min cast as xs:integer) and ($chartCOLUMNS le $sheetFILE//@max cast as xs:integer)
        let $chartAREAwdCUST := sum($colWIDTHS/@width/data() cast as xs:float)
    for $chartCOLUMNS in $chartFRcol to $chartTOcol
        for $colWIDTHS in $sheetFILE//col
        where ($chartCOLUMNS le $colWIDTHS//col/@min cast as xs:integer) or ($chartCOLUMNS ge $sheetFILE//@max cast as xs:integer)
        let $chartAREAwdDEF := count($colWIDTHS) * $colWIDTHdef
    return ($chartAREAwdCUST + $chartAREAwdDEF - $chartFOFFcol + $chartTOFFcol)
};

declare function local:chartHEIGHT()
    {for $chartNAMES in $sheetDWGSFILE//xdr:cNvPr[contains(@name, 'Chart')]
    let $chartFRrow as xs:integer := $chartNAMES/ancestor::xdr:graphicFrame/preceding-sibling::xdr:from/xdr:row/text() cast as xs:integer
    let $chartFOFFrow as xs:integer := $chartNAMES/ancestor::xdr:graphicFrame/preceding-sibling::xdr:from/xdr:rowoff/text() cast as xs:integer

    let $chartTOrow as xs:integer := $chartNAMES/ancestor::xdr:graphicFrame/preceding-sibling::xdr:to/xdr:row/text() cast as xs:integer
    let $chartTOFFrow as xs:integer := $chartNAMES/ancestor::xdr:graphicFrame/preceding-sibling::xdr:to/xdr:rowoff/text() cast as xs:integer 

    for $chartROWS in $chartFRrow to $chartTOrow
    (:the offset values below are in EMU and need conversion:)
    return (sum($chartROWS) *  $rowHEIGHTdef - $chartFOFFrow + $chartTOFFrow)
    };
    
element svgs
    {for $chartANCHORS in $sheetDWGSFILE//xdr:twoCellAnchor
    where $chartANCHORS//xdr:cNvPr[contains(@name, 'Chart')]
    let $chartFRrow := $chartANCHORS//xdr:from/xdr:row    
    let $chartFRcol := $chartANCHORS//xdr:from/xdr:col

    order by $chartFRrow, $chartFRcol

    let $chartrID := $chartANCHORS/xdr:graphicFrame//c:chart/@r:id/data()
    let $chartTGT := doc(concat(substring-before(document-uri(.),'workbook.xml'),'drawings/_rels',substring-after($sheetDWGSTGT,'drawings'),'.rels'))//rels:Relationship[@Id=$chartrID]/@Target/data() cast as xs:string
    let $chartFILE := doc(concat(substring-before(document-uri(.),'workbook.xml'),substring-after($chartTGT,'../')))

    let $plotXOrig := (try {($chartAREAwd - $plotAREAwd) div 2} catch * {$chartAREAwd * $chartFILE//plotArea//c:x/@val/data() cast as xs:float})
    let $plotYOrig := (try {($chartAREAht - $plotAREAht) div 2} catch * {$chartAREAht * $chartFILE//plotArea//c:y/@val/data() cast as xs:float})

    return 
    element svg{
    attribute width {concat($chartAREAwd,'cm')},
    attribute height {concat($chartAREAht,'cm')},
    attribute viewBox {concat('0 0 ',$chartAREAwd,' ',$chartAREAht)},
    attribute version {'1.1'},
    attribute xlink {'http://www.w3.org/1999/xlink'},
    
    let $chartTITLE := $chartFILE//c:chartSpace/c:chart/c:title//a:t/string()
    return local:svgTITLE($chartTITLE),

    for $AXES in $chartFILE//c:valAx
    where $AXES/c:delete/@val ne '1'
        let $axMAX := $AXES/c:scaling/c:max/@val/data()
        let $axMIN := $AXES/c:scaling/c:min/@val/data()
        let $axMINOR := $AXES//c:minorUnit/@val/data()
        let $axMAJOR := $AXES//c:majorUnit/@val/data()
        
        let $gridCOUNT := fn:round(($axMAX - $axMIN) div $axMINOR) cast as xs:integer
        let $gridMAX := if ($chartFILE//c:valAx//c:logBase) then 9 else $gridCOUNT - 1
        let $lineLABEL := if ($chartFILE//c:valAx//c:logBase) then $gridMAX else 2
        let $logSTEPS := if ($chartFILE//c:valAx//c:logBase) then math:log10($axMAX div $axMIN) cast as xs:integer else 1
        return(
        for $logSTEP in 1 to $logSTEPS
            for $gridLINE in 1 to $gridMAX
                let $VgridXcds := if ($chartFILE//c:valAx//c:logBase) 
                then ($plotXOrig + $plotAREAwd div $logSTEPS * (math:log10($axMIN * $gridLINE * math:exp10($logSTEP - 1)) - math:log10($axMIN)),$plotXOrig + $plotAREAwd div $logSTEPS * (math:log10($axMIN * $gridLINE * math:exp10($logSTEP - 1)) - math:log10($axMIN)))
                else ($plotXOrig + ($gridLINE div $gridCOUNT * $plotAREAwd),$plotXOrig + ($gridLINE div $gridCOUNT * $plotAREAwd))
                let $VgridYcds := ($plotYOrig,($plotYOrig + $plotAREAht))
                let $HgridXcds := ($plotXOrig,($plotXOrig + $plotAREAwd))
                let $HgridYcds := if ($chartFILE//c:valAx//c:logBase) 
                then ($plotYOrig + $plotAREAwd div $logSTEPS * (math:log10($axMIN * $gridLINE * math:exp10($logSTEP - 1)) - math:log10($axMIN)),$plotYOrig + $plotAREAwd div $logSTEPS * (math:log10($axMIN * $gridLINE * math:exp10($logSTEP - 1)) - math:log10($axMIN)))
                else ($plotYOrig + ($gridLINE div $gridCOUNT * $plotAREAht),$plotYOrig + ($gridLINE div $gridCOUNT * $plotAREAht))

                return (local:svgPATH(
                        concat('grid_',$AXES/c:axPos/@val,'_S',$logSTEP,'_L',$gridLINE),
                        $gridCOL,
                        try {$pathSTRK div 2} catch * {0.2},
                        switch ($AXES/c:axPos/@val)
                        case "b" return $VgridXcds
                        default return $HgridXcds,
                        switch ($AXES/c:axPos/@val)
                        case "b" return $VgridYcds
                        default return $HgridYcds
                        ),
                        if ($gridLINE mod $lineLABEL eq 0) then
                            local:svgTEXT(
                            concat('axlabel_',$AXES/c:axPos/@val,'_S',$logSTEP,'_L',$gridLINE),
                            switch ($AXES/c:axPos/@val)
                            case "b" return 'text-anchor: middle'
                            case "l" return 'text-anchor: end'
                            case "r" return 'text-anchor: start'
                            default return 'text-anchor: start',
                            $stdFONT,
                            try {$stdFONTptSZ} catch * {$AXES//c:txPr//@sz/data() div 100 cast as xs:decimal},
                            $axlabFONTcol,
                            switch ($AXES/c:axPos/@val)
                            case "b" return $VgridXcds[1]
                            case "l" return $HgridXcds[1]
                            case "r" return $HgridXcds[1] + $plotAREAwd
                            default return (),
                            switch ($AXES/c:axPos/@val)
                            case "b" return $VgridYcds[1]
                            default return $HgridYcds[1] + (try {$stdFONTptSZ} catch * {$AXES//c:txPr//@sz/data() div 100}) * $ptCMfactor div 2.75,
                            switch (boolean($chartFILE//c:valAx//c:logBase))
                            case false() return($axMIN  + ($gridLINE * $axMINOR)) cast as xs:float 
                            case true() return(math:exp10($logSTEP + math:log10($axMIN))) cast as xs:float
                            default return ()
                            )
                        else ()
                        ),
                    for $txtLINE at $lineNUM in $AXES/c:title//a:p
                        return local:svgTEXT(
                        concat('axtitle_',$AXES/c:axPos/@val, $lineNUM),
                        switch ($AXES/c:axPos/@val)
                        case "b" return 'text-anchor: middle'
                        case "l" return 'text-anchor: end'
                        case "r" return 'text-anchor: start'
                        default return 'text-anchor: start',
                        $stdFONT,
                        try {$stdFONTptSZ} catch * {$txtLINE[$lineNUM]//a:rPr[1]/@sz div 100 cast as xs:decimal},
                        $stdFONTcol,
                        $AXES/c:title/c:layout/c:manualLayout/c:x/@val * $chartAREAwd,
                        $AXES/c:title/c:layout/c:manualLayout/c:y/@val * $chartAREAht + ((try {$stdFONTptSZ} catch * {$txtLINE[$lineNUM]//a:rPr//@sz div 100 cast as xs:decimal}) * $ptCMfactor * ($lineNUM - 1)),
                        $txtLINE)
                 ),
                 for $scatterSERIES at $scatterIDX in $chartFILE//c:scatterChart
                    let $yValIDX := $scatterIDX * 2
                    return
                        for $pathSERIES at $pathIDX in $scatterSERIES/c:ser[not(descendant::c:ptCount[not(ancestor::c:strCache)]/@val = "1")]
                        let $pathTYPE := if ($pathSERIES//c:smooth/@val = 1) then 'smooth_' else 'straight_'
                        return local:svgPATH(
                        concat('PATH_',$pathTYPE,$scatterIDX),
                        try {$pathCOL} catch * {'black'},
                        $pathSTRK,
                        for $pathX in $pathSERIES/c:xVal//c:v
                        let $pathXcoords := if ($chartFILE//c:valAx//c:logBase)
                        then $plotXOrig + ((math:log10($pathX div $chartFILE//c:valAx[1]/c:scaling/c:min/@val/data()) * ($plotAREAwd div (math:log10($chartFILE//c:valAx[1]/c:scaling/c:max/@val/data() div $chartFILE//c:valAx[1]/c:scaling/c:min/@val/data())))))
                        else $plotXOrig + (($pathX - $chartFILE//c:valAx[1]/c:scaling/c:min/@val/data()) * ($plotAREAwd div ($chartFILE//c:valAx[1]/c:scaling/c:max/@val/data() - $chartFILE//c:valAx[1]/c:scaling/c:min/@val/data())))
                        return $pathXcoords,
                        for $pathY in $pathSERIES/c:yVal//c:v
                        let $pathYcoords :=  if ($chartFILE//c:valAx//c:logBase)
                        then $plotYOrig + ((math:log10($pathY div $chartFILE//c:valAx[$yValIDX]/c:scaling/c:min/@val/data()) * ($plotAREAht div (math:log10($chartFILE//c:valAx[$yValIDX]/c:scaling/c:max/@val/data() div $chartFILE//c:valAx[$yValIDX]/c:scaling/c:min/@val/data())))))
                        else $plotYOrig + (($pathY - $chartFILE//c:valAx[$yValIDX]/c:scaling/c:min/@val/data()) * ($plotAREAht div ($chartFILE//c:valAx[$yValIDX]/c:scaling/c:max/@val/data() - $chartFILE//c:valAx[$yValIDX]/c:scaling/c:min/@val/data())))
                        return $pathYcoords
               ),
    for $plotDIMS in $chartFILE//c:plotArea
    return local:svgRECT(
        (try {($chartAREAwd - $plotAREAwd) div 2} catch * {$chartAREAwd * $chartFILE//plotArea//c:x/@val/data()}),
        (try {($chartAREAht - $plotAREAht) div 2} catch * {$chartAREAht * $chartFILE//plotArea//c:y/@val/data()}),
        (try {$plotAREAwd} catch * {$chartAREAwd * $chartFILE//plotArea//c:w/@val/data()}),
        (try {$plotAREAht} catch * {$chartAREAht * $chartFILE//plotArea//c:h/@val/data()}),
        'none',
        'black',
        $pathSTRK),
(:find drawings file via relationships and retrieve text box info:)
        let $UshapesID := $chartFILE//c:userShapes/@r:id/data()
        let $chartTXT := substring-after(document-uri($chartFILE),'/xl/charts/')
        let $chartRELS := concat(substring-before(document-uri($chartFILE),$chartTXT),'_rels/',$chartTXT,'.rels') cast as xs:anyURI
        
        where doc-available($chartRELS) eq true()
    
        let $chartDWGS := concat(substring-before(document-uri($chartFILE),'charts/'),substring-after(doc($chartRELS)/rels:Relationships/rels:Relationship[@Id=$UshapesID]/@Target/data(),'../')) cast as xs:anyURI   
        for $txtBOX at $txtBOXnum in doc($chartDWGS)//cdr:relSizeAnchor[*//a:r/a:t]
        where $txtBOX//@txBox eq '1'

        let $txtBOXx0 := $txtBOX/cdr:from/cdr:x * $chartAREAwd
        let $txtBOXy0 := $txtBOX/cdr:from/cdr:y * $chartAREAht
        let $txtBOXx1 := $txtBOX/cdr:to/cdr:x * $chartAREAwd - $txtBOXx0
        let $txtBOXy1 := $txtBOX/cdr:to/cdr:y * $chartAREAht - $txtBOXy0
        return (local:svgRECT(
                $txtBOXx0,
                $txtBOXy0,
                $txtBOXx1,
                $txtBOXy1,
                'white',
                'black',
                0),
            let $txtPOSx := $txtBOXx0 + $txtBOXx1 div 2
            let $txtPOSy := $txtBOXy0 + $txtBOXy1 div 2

            for $txtLINE at $lineNUM in $txtBOX//a:p
                return local:svgTEXT(
                concat('txtbox_',$txtBOXnum),
                'text-anchor: middle',
                $stdFONT,
                try {$stdFONTptSZ} catch * {$txtLINE[$lineNUM]//a:rPr[1]/@sz div 100 cast as xs:decimal},
                $stdFONTcol,
                $txtPOSx,
                $txtPOSy + ((try {$stdFONTptSZ} catch * {$txtLINE[$lineNUM]//a:rPr//@sz div 100 cast as xs:decimal}) * $ptCMfactor * ($lineNUM - 1)),
                $txtLINE)
         )
    }
}
