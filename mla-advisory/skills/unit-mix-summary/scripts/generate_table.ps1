# generate_table.ps1
# Generates a formatted Unit Mix Summary Word document from JSON data
# Usage: .\generate_table.ps1 -DataJson '<json>' -OutputPath 'Unit Mix Summary.docx'
#
# All Word COM measurements are in POINTS (1 inch = 72 pts)

param(
    [Parameter(Mandatory=$true)]
    [string]$DataJson,

    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "Unit Mix Summary.docx"
)

$data     = $DataJson | ConvertFrom-Json
$rows     = $data.rows

# Resolve to absolute path
if (-not [System.IO.Path]::IsPathRooted($OutputPath)) {
    $OutputPath = Join-Path (Get-Location) $OutputPath
}

# --- Word color: BGR integer from RRGGBB hex ---
function ConvertTo-WordColor([string]$hex) {
    $r = [Convert]::ToInt32($hex.Substring(0,2), 16)
    $g = [Convert]::ToInt32($hex.Substring(2,2), 16)
    $b = [Convert]::ToInt32($hex.Substring(4,2), 16)
    return $b * 65536 + $g * 256 + $r
}

$colorDark    = ConvertTo-WordColor "1A1A1A"   # near-black header
$colorWhite   = ConvertTo-WordColor "FFFFFF"
$colorYellow  = ConvertTo-WordColor "FFFF00"   # unit type column
$colorAltRow  = ConvertTo-WordColor "F2F2F2"   # alternating row

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible       = $false
    $word.DisplayAlerts = 0

    $doc = $word.Documents.Add()

    # --- Page setup: landscape, 0.75 in margins (54 pts each) ---
    $section = $doc.Sections.Item(1)
    $section.PageSetup.Orientation   = 1    # wdOrientLandscape
    $section.PageSetup.LeftMargin    = 54
    $section.PageSetup.RightMargin   = 54
    $section.PageSetup.TopMargin     = 54
    $section.PageSetup.BottomMargin  = 54

    # --- Table ---
    $headers   = @('UNIT TYPE','%-INVENTORY','MIN-SIZE (SF)','MAX-SIZE (SF)','AVERAGE-SIZE','MEDIAN-SIZE')
    # Landscape page = 792 pts wide. Available = 792 - 108 = 684 pts.
    # Col widths in pts: [130, 111, 111, 111, 111, 111] = 685 — snug fit
    $colWidths = @(130, 111, 111, 111, 111, 110)

    $numRows = $rows.Count + 1  # +1 for header row

    $range = $doc.Content
    $range.Collapse(0)  # wdCollapseEnd
    $table = $doc.Tables.Add($range, $numRows, 6)
    $table.Style = "Table Grid"
    $table.Borders.InsideLineStyle  = 1  # wdLineStyleSingle
    $table.Borders.OutsideLineStyle = 1

    # Set column widths
    for ($c = 1; $c -le 6; $c++) {
        $table.Columns.Item($c).Width = $colWidths[$c - 1]
    }

    # --- Header row ---
    $rowHeight = 18  # 0.25 in in points

    $headerRow = $table.Rows.Item(1)
    $headerRow.HeightRule = 1   # wdRowHeightAtLeast
    $headerRow.Height = $rowHeight

    for ($c = 1; $c -le 6; $c++) {
        $cell = $table.Cell(1, $c)
        $cell.Shading.BackgroundPatternColor = $colorDark
        $cell.VerticalAlignment = 1   # wdCellAlignVerticalCenter

        $cr = $cell.Range
        $cr.Text = $headers[$c - 1]
        $cr.ParagraphFormat.Alignment = 1   # wdAlignParagraphCenter
        $cr.Font.Name  = "Calibri"
        $cr.Font.Size  = 8.5
        $cr.Font.Bold  = $true
        $cr.Font.Color = $colorWhite
    }

    # --- Data rows ---
    for ($r = 0; $r -lt $rows.Count; $r++) {
        $row     = $rows[$r]
        $wordRow = $table.Rows.Item($r + 2)
        $wordRow.HeightRule = 1
        $wordRow.Height = $rowHeight

        $values = @(
            $row.unit_type,
            $row.pct_inventory,
            [string]$row.min_size,
            [string]$row.max_size,
            [string]$row.avg_size,
            [string]$row.median_size
        )

        $altBg = if ($r % 2 -eq 1) { $colorAltRow } else { $colorWhite }

        for ($c = 1; $c -le 6; $c++) {
            $cell = $table.Cell($r + 2, $c)
            $cell.VerticalAlignment = 1

            $cell.Shading.BackgroundPatternColor = if ($c -eq 1) { $colorYellow } else { $altBg }

            $cr = $cell.Range
            $cr.Text = $values[$c - 1]
            $cr.ParagraphFormat.Alignment = 1
            $cr.Font.Name  = "Calibri"
            $cr.Font.Size  = 9
            $cr.Font.Bold  = ($c -eq 1)
            $cr.Font.Color = $colorDark
        }
    }

    # Save and close (SaveAs2 is more reliable via COM)
    $doc.SaveAs2($OutputPath, 16)  # 16 = wdFormatDocx
    $doc.Close()
    $word.Quit()

    Write-Output "Saved: $OutputPath"

} catch {
    try { $doc.Close($false) }  catch {}
    try { $word.Quit() }        catch {}
    Write-Error "Error generating document: $_"
    exit 1
} finally {
    if ($word) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    [System.GC]::Collect()
}
