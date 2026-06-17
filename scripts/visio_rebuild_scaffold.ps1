param(
    [Parameter(Mandatory=$true)]
    [string]$VsdxPath,

    [double]$PageW = 16.0,
    [double]$PageH = 12.0,
    [double]$RefW = 1448.0,
    [double]$RefH = 1086.0,

    [string]$PreviewPath,

    [string[]]$ExportFormats,

    [string]$OutputDir,
    [string]$OutputBaseName
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'visio_export_formats.ps1')

function VX([double]$x) { $script:PageW * $x / $script:RefW }
function VY([double]$y) { $script:PageH - ($script:PageH * $y / $script:RefH) }
function RGBF([int]$r, [int]$g, [int]$b) { "RGB($r,$g,$b)" }

$C = @{
    Blue = RGBF 31 95 184
    Purple = RGBF 122 88 166
    Green = RGBF 90 132 64
    Teal = RGBF 20 132 150
    Orange = RGBF 231 116 26
    Black = RGBF 17 17 17
    Gray = RGBF 95 95 95
    White = RGBF 255 255 255
    BlueSoft = RGBF 243 248 255
    PurpleSoft = RGBF 251 248 255
    GreenSoft = RGBF 246 251 241
    OrangeSoft = RGBF 255 248 239
}

function Set-Cell($shape, [string]$cell, [string]$formula) {
    try { $shape.CellsU($cell).FormulaU = $formula } catch {}
}

function Style-Shape($shape, [string]$fill, [string]$line, [double]$linePt = 0.8, [int]$dash = 1, [double]$roundPx = 0) {
    if ($fill -eq 'none') {
        Set-Cell $shape 'FillPattern' '0'
    } else {
        Set-Cell $shape 'FillPattern' '1'
        Set-Cell $shape 'FillForegnd' $fill
    }
    if ($line -eq 'none') {
        Set-Cell $shape 'LinePattern' '0'
    } else {
        Set-Cell $shape 'LinePattern' ([string]$dash)
        Set-Cell $shape 'LineColor' $line
        Set-Cell $shape 'LineWeight' "$linePt pt"
    }
    if ($roundPx -gt 0) {
        Set-Cell $shape 'Rounding' ((VX $roundPx).ToString([Globalization.CultureInfo]::InvariantCulture) + ' in')
    }
}

function Set-Text($shape, [string]$text, [double]$size = 10, [string]$color = $C.Black, [bool]$bold = $false, [bool]$italic = $false, [int]$align = 1) {
    $shape.Text = $text
    Set-Cell $shape 'Char.Font' '0'
    Set-Cell $shape 'Char.Size' "$size pt"
    Set-Cell $shape 'Char.Color' $color
    $style = 0
    if ($bold) { $style += 1 }
    if ($italic) { $style += 2 }
    Set-Cell $shape 'Char.Style' ([string]$style)
    Set-Cell $shape 'Para.HorzAlign' ([string]$align)
    Set-Cell $shape 'VerticalAlign' '1'
    foreach ($m in 'TxtMarginLeft','TxtMarginRight','TxtMarginTop','TxtMarginBottom') {
        Set-Cell $shape $m '1 pt'
    }
}

function RectTL([double]$x, [double]$y, [double]$w, [double]$h, [string]$text = '', [string]$fill = 'none', [string]$line = $C.Black, [double]$size = 10, [bool]$bold = $false, [double]$linePt = 0.8, [int]$dash = 1, [double]$roundPx = 6) {
    $s = $script:Page.DrawRectangle((VX $x), (VY ($y + $h)), (VX ($x + $w)), (VY $y))
    Style-Shape $s $fill $line $linePt $dash $roundPx
    if ($text -ne '') { Set-Text $s $text $size $C.Black $bold }
    return $s
}

function TextTL([double]$x, [double]$y, [double]$w, [double]$h, [string]$text, [double]$size = 10, [string]$color = $C.Black, [bool]$bold = $false, [bool]$italic = $false, [int]$align = 1) {
    $s = RectTL $x $y $w $h '' 'none' 'none' $size $bold 0 1 0
    Set-Text $s $text $size $color $bold $italic $align
    return $s
}

function OvalTL([double]$x, [double]$y, [double]$w, [double]$h, [string]$text = '', [string]$fill = $C.White, [string]$line = $C.Black, [double]$size = 8, [bool]$bold = $false, [double]$linePt = 0.8) {
    $s = $script:Page.DrawOval((VX $x), (VY ($y + $h)), (VX ($x + $w)), (VY $y))
    Style-Shape $s $fill $line $linePt 1 0
    if ($text -ne '') { Set-Text $s $text $size $C.Black $bold }
    return $s
}

function LineTL([double]$x1, [double]$y1, [double]$x2, [double]$y2, [string]$color = $C.Black, [double]$linePt = 0.8, [bool]$arrowEnd = $false, [bool]$arrowBegin = $false, [int]$dash = 1) {
    $s = $script:Page.DrawLine((VX $x1), (VY $y1), (VX $x2), (VY $y2))
    Set-Cell $s 'LineColor' $color
    Set-Cell $s 'LineWeight' "$linePt pt"
    Set-Cell $s 'LinePattern' ([string]$dash)
    if ($arrowEnd) { Set-Cell $s 'EndArrow' '4' }
    if ($arrowBegin) { Set-Cell $s 'BeginArrow' '4' }
    return $s
}

function DotTL([double]$cx, [double]$cy, [double]$r, [string]$fill, [string]$line = $C.White) {
    return OvalTL ($cx - $r) ($cy - $r) (2 * $r) (2 * $r) '' $fill $line 6 $false 0.4
}

function Assert-RelBox([double]$u, [double]$v, [double]$uw, [double]$vh, [string]$label = 'relative box') {
    if ($u -lt 0 -or $v -lt 0 -or $uw -lt 0 -or $vh -lt 0 -or ($u + $uw) -gt 1 -or ($v + $vh) -gt 1) {
        throw "$label is outside calibrated panel bounds. Use 0-1 local coordinates or enlarge the parent panel."
    }
}

function Assert-RelPoint([double]$u, [double]$v, [string]$label = 'relative point') {
    if ($u -lt 0 -or $v -lt 0 -or $u -gt 1 -or $v -gt 1) {
        throw "$label is outside calibrated panel bounds. Use 0-1 local coordinates or enlarge the parent panel."
    }
}

function RX([double]$x0, [double]$w0, [double]$u) { $x0 + $w0 * $u }
function RY([double]$y0, [double]$h0, [double]$v) { $y0 + $h0 * $v }

function RectRel([double]$x0, [double]$y0, [double]$w0, [double]$h0, [double]$u, [double]$v, [double]$uw, [double]$vh, [string]$text = '', [string]$fill = 'none', [string]$line = $C.Black, [double]$size = 10, [bool]$bold = $false, [double]$linePt = 0.8, [int]$dash = 1, [double]$roundPx = 6) {
    Assert-RelBox $u $v $uw $vh $text
    return RectTL (RX $x0 $w0 $u) (RY $y0 $h0 $v) ($w0 * $uw) ($h0 * $vh) $text $fill $line $size $bold $linePt $dash $roundPx
}

function TextRel([double]$x0, [double]$y0, [double]$w0, [double]$h0, [double]$u, [double]$v, [double]$uw, [double]$vh, [string]$text, [double]$size = 10, [string]$color = $C.Black, [bool]$bold = $false, [bool]$italic = $false, [int]$align = 1) {
    Assert-RelBox $u $v $uw $vh $text
    return TextTL (RX $x0 $w0 $u) (RY $y0 $h0 $v) ($w0 * $uw) ($h0 * $vh) $text $size $color $bold $italic $align
}

function OvalRel([double]$x0, [double]$y0, [double]$w0, [double]$h0, [double]$u, [double]$v, [double]$uw, [double]$vh, [string]$text = '', [string]$fill = $C.White, [string]$line = $C.Black, [double]$size = 8, [bool]$bold = $false, [double]$linePt = 0.8) {
    Assert-RelBox $u $v $uw $vh $text
    return OvalTL (RX $x0 $w0 $u) (RY $y0 $h0 $v) ($w0 * $uw) ($h0 * $vh) $text $fill $line $size $bold $linePt
}

function LineRel([double]$x0, [double]$y0, [double]$w0, [double]$h0, [double]$u1, [double]$v1, [double]$u2, [double]$v2, [string]$color = $C.Black, [double]$linePt = 0.8, [bool]$arrowEnd = $false, [bool]$arrowBegin = $false, [int]$dash = 1) {
    Assert-RelPoint $u1 $v1 'line start'
    Assert-RelPoint $u2 $v2 'line end'
    return LineTL (RX $x0 $w0 $u1) (RY $y0 $h0 $v1) (RX $x0 $w0 $u2) (RY $y0 $h0 $v2) $color $linePt $arrowEnd $arrowBegin $dash
}

function Draw-ReferenceFigure {
    # Replace this with the task-specific drawing code.
    # Keep the order: panels -> main flow -> text boxes -> repeated motifs -> annotations.
    # For complex panels, calibrate the panel bounds and draw internals with RectRel/TextRel/LineRel.
    RectTL 20 60 180 220 'Input Sequence' $C.BlueSoft $C.Blue 11 $true 1.0 1 8 | Out-Null
    $panelX = 260.0
    $panelY = 60.0
    $panelW = 220.0
    $panelH = 220.0
    RectTL $panelX $panelY $panelW $panelH 'Block 1' $C.White $C.Blue 10 $true 1.0 1 8 | Out-Null
    RectRel $panelX $panelY $panelW $panelH 0.14 0.25 0.72 0.16 'Module A' $C.PurpleSoft $C.Purple 11 $true 0.8 1 5 | Out-Null
    LineTL 200 170 260 170 $C.Black 1.0 $true | Out-Null
    LineTL 480 170 570 170 $C.Black 1.0 $true | Out-Null
    RectTL 570 90 250 160 'Processing' $C.GreenSoft $C.Green 11 $true 1.0 1 8 | Out-Null
    LineTL 820 170 880 170 $C.Black 1.0 $true | Out-Null
    RectTL 880 90 240 160 'Output' $C.OrangeSoft $C.Orange 11 $true 1.0 1 8 | Out-Null
    TextTL 600 20 360 28 'Repeated Processing Stage' 13 $C.Blue $true | Out-Null
}

$backup = Join-Path (Split-Path -Parent $VsdxPath) (([IO.Path]::GetFileNameWithoutExtension($VsdxPath)) + ".backup-" + (Get-Date -Format 'yyyyMMdd-HHmmss') + ".vsdx")
Copy-Item -LiteralPath $VsdxPath -Destination $backup
Write-Output "Backup: $backup"

$visio = $null
$doc = $null
try {
    $visio = New-Object -ComObject Visio.Application
    $visio.Visible = $true
    $doc = $visio.Documents.Open($VsdxPath)
    $script:Page = $doc.Pages.Item(1)
    $script:PageW = $PageW
    $script:PageH = $PageH
    $script:RefW = $RefW
    $script:RefH = $RefH

    $script:Page.PageSheet.CellsU('PageWidth').FormulaU = "$PageW in"
    $script:Page.PageSheet.CellsU('PageHeight').FormulaU = "$PageH in"
    while ($script:Page.Shapes.Count -gt 0) {
        $script:Page.Shapes.Item(1).Delete() | Out-Null
    }

    Draw-ReferenceFigure

    $doc.Save() | Out-Null

    $formatsToExport = New-Object System.Collections.Generic.List[string]
    if ($PreviewPath -and -not $formatsToExport.Contains('png')) {
        $formatsToExport.Add('png') | Out-Null
    }
    foreach ($format in @($ExportFormats)) {
        if ($format -and -not $formatsToExport.Contains($format.ToLowerInvariant())) {
            $formatsToExport.Add($format.ToLowerInvariant()) | Out-Null
        }
    }
    if ($formatsToExport.Count -gt 0) {
        Export-VisioPageFormats `
            -Document $doc `
            -Page $script:Page `
            -SourcePath $VsdxPath `
            -Formats @($formatsToExport) `
            -OutputDir $OutputDir `
            -OutputBaseName $OutputBaseName `
            -PreviewPath $PreviewPath
    }

    Write-Output "Saved: $VsdxPath"
} finally {
    if ($doc -ne $null) {
        try { $doc.Close() } catch {}
    }
    if ($visio -ne $null) {
        try { $visio.Quit() } catch {}
    }
}
