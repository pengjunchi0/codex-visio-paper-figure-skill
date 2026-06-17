$script:VisioExportSupportedFormats = @('png', 'svg', 'pdf', 'pptx')

function Normalize-VisioExportFormats {
    param([string[]]$Formats)

    $normalized = New-Object System.Collections.Generic.List[string]
    foreach ($formatGroup in $Formats) {
        if (-not $formatGroup) { continue }
        foreach ($format in ($formatGroup -split ',')) {
            if (-not $format) { continue }
            $name = $format.Trim().TrimStart('.').ToLowerInvariant()
            if ($script:VisioExportSupportedFormats -notcontains $name) {
                throw "Unsupported export format '$format'. Supported formats: $($script:VisioExportSupportedFormats -join ', ')"
            }
            if (-not $normalized.Contains($name)) {
                $normalized.Add($name)
            }
        }
    }
    return @($normalized)
}

function Resolve-VisioExportPath {
    param(
        [Parameter(Mandatory=$true)]
        [string]$SourcePath,

        [Parameter(Mandatory=$true)]
        [string]$Format,

        [string]$OutputDir,
        [string]$OutputBaseName,
        [string]$PreviewPath
    )

    if ($Format -eq 'png' -and $PreviewPath) {
        return [IO.Path]::GetFullPath($PreviewPath)
    }

    if (-not $OutputDir) {
        $OutputDir = Split-Path -Parent $SourcePath
    }
    if (-not $OutputBaseName) {
        $OutputBaseName = [IO.Path]::GetFileNameWithoutExtension($SourcePath)
    }
    if (-not (Test-Path -LiteralPath $OutputDir)) {
        New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
    }

    return [IO.Path]::GetFullPath((Join-Path $OutputDir "$OutputBaseName.$Format"))
}

function Export-VisioPdf {
    param(
        [Parameter(Mandatory=$true)]
        $Document,

        [Parameter(Mandatory=$true)]
        [string]$OutPath
    )

    $dir = Split-Path -Parent $OutPath
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
    }

    try {
        # 1 = PDF, 1 = print quality, 0 = all pages. Numeric constants avoid requiring Visio interop assemblies.
        $Document.ExportAsFixedFormat(1, $OutPath, 1, 0)
    } catch {
        throw "PDF export failed: $($_.Exception.Message)"
    }
}

function Export-VisioPptx {
    param(
        [Parameter(Mandatory=$true)]
        $Page,

        [Parameter(Mandatory=$true)]
        [string]$OutPath
    )

    $dir = Split-Path -Parent $OutPath
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
    }

    $scratch = Join-Path ([IO.Path]::GetTempPath()) ("visio-export-" + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Force -Path $scratch | Out-Null
    $svgPath = Join-Path $scratch 'page.svg'

    $powerPoint = $null
    $presentation = $null
    try {
        $Page.Export($svgPath)
        $pageWidthPt = [double]$Page.PageSheet.CellsU('PageWidth').ResultIU * 72.0
        $pageHeightPt = [double]$Page.PageSheet.CellsU('PageHeight').ResultIU * 72.0

        try {
            $powerPoint = New-Object -ComObject PowerPoint.Application
        } catch {
            throw "PowerPoint COM is required for PPTX export: $($_.Exception.Message)"
        }

        $powerPoint.Visible = $true
        $presentation = $powerPoint.Presentations.Add()
        $presentation.PageSetup.SlideWidth = $pageWidthPt
        $presentation.PageSetup.SlideHeight = $pageHeightPt

        # 12 = ppLayoutBlank, 24 = ppSaveAsOpenXMLPresentation.
        $slide = $presentation.Slides.Add(1, 12)
        $slide.Shapes.AddPicture($svgPath, $false, $true, 0, 0, $pageWidthPt, $pageHeightPt) | Out-Null
        $presentation.SaveAs($OutPath, 24)
    } catch {
        throw "PPTX export failed: $($_.Exception.Message)"
    } finally {
        if ($presentation -ne $null) {
            try { $presentation.Close() } catch {}
        }
        if ($powerPoint -ne $null) {
            try { $powerPoint.Quit() } catch {}
        }
        if (Test-Path -LiteralPath $scratch) {
            Remove-Item -LiteralPath $scratch -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Export-VisioPageFormats {
    param(
        [Parameter(Mandatory=$true)]
        $Document,

        [Parameter(Mandatory=$true)]
        $Page,

        [Parameter(Mandatory=$true)]
        [string]$SourcePath,

        [string[]]$Formats = @('png'),
        [string]$OutputDir,
        [string]$OutputBaseName,
        [string]$PreviewPath
    )

    $formatsToExport = Normalize-VisioExportFormats $Formats
    foreach ($format in $formatsToExport) {
        $outPath = Resolve-VisioExportPath -SourcePath $SourcePath -Format $format -OutputDir $OutputDir -OutputBaseName $OutputBaseName -PreviewPath $PreviewPath
        switch ($format) {
            'png' { $Page.Export($outPath) }
            'svg' { $Page.Export($outPath) }
            'pdf' { Export-VisioPdf -Document $Document -OutPath $outPath }
            'pptx' { Export-VisioPptx -Page $Page -OutPath $outPath }
        }

        $bytes = (Get-Item -LiteralPath $outPath).Length
        Write-Output ("{0}: {1} ({2} bytes)" -f $format.ToUpperInvariant(), $outPath, $bytes)
    }
}

function Export-VisioDocumentFormats {
    param(
        [Parameter(Mandatory=$true)]
        [string]$VsdxPath,

        [string[]]$Formats = @('png'),
        [string]$OutputDir,
        [string]$OutputBaseName,
        [int]$PageIndex = 1,
        [string]$PreviewPath,
        [switch]$Visible
    )

    $visio = $null
    $doc = $null
    try {
        $visio = New-Object -ComObject Visio.Application
        $visio.Visible = [bool]$Visible
        $doc = $visio.Documents.Open($VsdxPath)
        $page = $doc.Pages.Item($PageIndex)
        Export-VisioPageFormats -Document $doc -Page $page -SourcePath $VsdxPath -Formats $Formats -OutputDir $OutputDir -OutputBaseName $OutputBaseName -PreviewPath $PreviewPath
    } finally {
        if ($doc -ne $null) {
            try { $doc.Close() } catch {}
        }
        if ($visio -ne $null) {
            try { $visio.Quit() } catch {}
        }
    }
}
