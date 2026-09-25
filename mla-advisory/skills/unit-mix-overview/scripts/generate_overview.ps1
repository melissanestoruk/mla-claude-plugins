# MLA Unit Mix Overview page: N transposed tables (plan types as columns,
# INVENTORY/MIN/MAX/AVERAGE/MEDIAN rows) rendered via Word COM, landscape.
# Input: the render JSON produced by aggregate_overview.py.
param(
    [Parameter(Mandatory=$true)] [string]$DataPath,
    [Parameter(Mandatory=$true)] [string]$OutputPath
)

$data = Get-Content $DataPath -Raw | ConvertFrom-Json

function ConvertTo-WordColor([string]$hex) {
    $r = [Convert]::ToInt32($hex.Substring(0,2), 16)
    $g = [Convert]::ToInt32($hex.Substring(2,2), 16)
    $b = [Convert]::ToInt32($hex.Substring(4,2), 16)
    return $b * 65536 + $g * 256 + $r
}

$colorDark   = ConvertTo-WordColor "1A1A1A"
$colorWhite  = ConvertTo-WordColor "FFFFFF"
$colorAltRow = ConvertTo-WordColor "F2F2F2"
$colorBlue   = ConvertTo-WordColor "D9E2F3"

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Add()

    $section = $doc.Sections.Item(1)
    $section.PageSetup.Orientation  = 1   # landscape
    $section.PageSetup.LeftMargin   = 54
    $section.PageSetup.RightMargin  = 54
    $section.PageSetup.TopMargin    = 54
    $section.PageSetup.BottomMargin = 54

    $sel = $word.Selection

    $sel.Font.Name = "Butler"
    $sel.Font.Size = 22
    $sel.Font.Bold = $true
    $sel.Font.Color = $colorDark
    $sel.TypeText($data.title)
    $sel.TypeParagraph()

    $sel.Font.Name = "Montserrat Light"
    $sel.Font.Size = 9.5
    $sel.Font.Bold = $false
    $sel.TypeText($data.intro)
    $sel.TypeParagraph()

    $rowLabels = @('PLAN TYPE','INVENTORY','MIN','MAX','AVERAGE','MEDIAN')
    $fields    = @($null,'pct','min','max','avg','med')

    foreach ($sec in $data.sections) {
        $sel.EndKey(6) | Out-Null
        $sel.TypeParagraph()
        $sel.Font.Name = "Montserrat"
        $sel.Font.Size = 10
        $sel.Font.Bold = $true
        $sel.Font.Color = $colorDark
        $sel.TypeText($sec.label)
        $sel.TypeParagraph()

        $sel.Font.Name = "Montserrat Light"
        $sel.Font.Size = 9.5
        $sel.Font.Bold = $false
        $sel.TypeText($sec.blurb)
        $sel.TypeParagraph()

        $cols = $sec.columns
        $numCols = $cols.Count + 1
        $range = $doc.Content
        $range.Collapse(0)
        $table = $doc.Tables.Add($range, 6, $numCols)
        $table.Style = "Table Grid"
        $table.Borders.InsideLineStyle  = 1
        $table.Borders.OutsideLineStyle = 1

        # 684 pts printable width; label column 90, type columns split the rest
        $table.Columns.Item(1).Width = 90
        $typeColWidth = [Math]::Floor((684 - 90) / $cols.Count)
        for ($c = 2; $c -le $numCols; $c++) { $table.Columns.Item($c).Width = $typeColWidth }

        for ($r = 1; $r -le 6; $r++) {
            $wordRow = $table.Rows.Item($r)
            $wordRow.HeightRule = 1
            $wordRow.Height = 18

            # header white, stat rows banded F2F2F2/white, MEDIAN row light blue
            $rowBg = $colorWhite
            if ($r -eq 2 -or $r -eq 4) { $rowBg = $colorAltRow }
            if ($r -eq 6) { $rowBg = $colorBlue }

            for ($c = 1; $c -le $numCols; $c++) {
                $cell = $table.Cell($r, $c)
                $cell.VerticalAlignment = 1
                # label column matches its row's shading - never highlighted
                $cell.Shading.BackgroundPatternColor = if ($r -eq 1) { $colorWhite } else { $rowBg }

                if ($c -eq 1) { $text = $rowLabels[$r - 1] }
                elseif ($r -eq 1) { $text = $cols[$c - 2].type }
                else { $text = [string]$cols[$c - 2].($fields[$r - 1]) }

                $cr = $cell.Range
                $cr.Text = $text
                $cr.ParagraphFormat.Alignment = 1
                $cr.Font.Name  = "Montserrat"
                $cr.Font.Size  = 8
                $cr.Font.Bold  = ($r -eq 1 -or $c -eq 1)
                $cr.Font.Color = $colorDark
            }
        }

        $end = $doc.Content
        $end.Collapse(0)
        $sel.EndKey(6) | Out-Null
        $sel.Font.Name = "Montserrat Light"
        $sel.Font.Size = 7.5
        $sel.Font.Bold = $false
        $sel.Font.Color = $colorDark
        $sel.TypeText($sec.footnote)
        $sel.TypeParagraph()
    }

    $doc.SaveAs2($OutputPath, 16)
    $doc.Close()
    $word.Quit()
    Write-Output "Saved: $OutputPath"
} catch {
    try { $doc.Close($false) } catch {}
    try { $word.Quit() } catch {}
    Write-Error "Error generating document: $_"
    exit 1
} finally {
    if ($word) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null }
    [System.GC]::Collect()
}
