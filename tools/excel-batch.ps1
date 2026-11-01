# Apre con Excel vero tutti gli .xlsx di una cartella (una sola istanza privata e nascosta) e scrive un JSON per file:
# esito apertura, ripristino/errore, celle (Value2/Formula invarianti), formati, unioni, larghezze, oggetti (grafici, tabelle...).
# Mai FormulaLocal/Text. Il rilevatore coglie sia gli errori sia il ripristino silenzioso ("[Ripristinato]/[Repaired]" nel nome).
param([Parameter(Mandatory)][string]$InDir, [Parameter(Mandatory)][string]$OutDir, [switch]$Recalc)
. "$PSScriptRoot\Excel-Com.ps1"
$en = [System.Globalization.CultureInfo]::GetCultureInfo('en-US')
function GetP($o, [string]$n) { [System.__ComObject].InvokeMember($n, [System.Reflection.BindingFlags]::GetProperty, $null, $o, $null, $en) }
$null = New-Item -ItemType Directory -Force $OutDir
$h = Start-Xl
try {
    $xl = $h.App
    foreach ($f in Get-ChildItem $InDir -Filter *.xlsx) {
        $res = [ordered]@{ file = $f.Name; ok = $false; repaired = $false; error = $null; name = $null; date1904 = $null; names = @(); sheets = @() }
        $wb = $null
        try {
            try { $wb = $xl.Workbooks.Open($f.FullName, 0, $true) } catch { $res.error = $_.Exception.Message }
            if ($wb) {
                $res.name = $wb.Name
                $res.repaired = [bool]($wb.Name -match '\[(Repaired|Ripristinat\w*)\]')
                $res.date1904 = [bool]$wb.Date1904
                if ($Recalc) { $xl.CalculateFull() }
                foreach ($nm in $wb.Names) { $res.names += [ordered]@{ n = $nm.Name; r = $nm.RefersTo } }
                foreach ($ws in $wb.Worksheets) {
                    $ur = $ws.UsedRange
                    $r0 = $ur.Row; $c0 = $ur.Column; $nr = $ur.Rows.Count; $nc = $ur.Columns.Count
                    $cells = New-Object System.Collections.ArrayList
                    $v2 = $ur.Value2; $fm = $ur.Formula
                    $styleBudget = 600
                    for ($i = 1; $i -le $nr; $i++) { for ($j = 1; $j -le $nc; $j++) {
                        if ($nr -eq 1 -and $nc -eq 1) { $v = $v2; $fo = $fm } else { $v = $v2[$i, $j]; $fo = $fm[$i, $j] }
                        if ($null -eq $v -and ($fo -eq '' -or $null -eq $fo)) { continue }
                        $t = 'n'; $e = $null
                        if ($v -is [int]) { $t = 'e'; $e = $v; $v = $null } elseif ($v -is [string]) { $t = 's' } elseif ($v -is [bool]) { $t = 'b' }
                        $cell = [ordered]@{ r = $r0 + $i - 1; c = $c0 + $j - 1; t = $t; v = $v; e = $e; f = $(if (($fo -is [string]) -and $fo.StartsWith('=')) { $fo } else { $null }) }
                        if ($styleBudget -gt 0) {
                            $styleBudget--
                            $c = $ws.Cells.Item($r0 + $i - 1, $c0 + $j - 1)
                            $cell.nf = GetP $c 'NumberFormat'; $cell.b = [bool]$c.Font.Bold; $cell.i = [bool]$c.Font.Italic; $cell.fc = $c.Font.Color; $cell.fill = $(if ($c.Interior.ColorIndex -eq -4142) { $null } else { $c.Interior.Color })
                            $cell.ha = $c.HorizontalAlignment; $cell.wrap = $c.WrapText
                            if ($c.MergeCells -and $c.MergeArea.Cells.Item(1, 1).Address() -eq $c.Address()) { $cell.merge = $c.MergeArea.Address($false, $false) }
                        }
                        [void]$cells.Add($cell)
                    } }
                    $cw = @(); for ($k = 1; $k -le [Math]::Min(12, $c0 + $nc); $k++) { $cw += [Math]::Round([double]$ws.Columns.Item($k).ColumnWidth, 3) }
                    $rh = @(); for ($k = 1; $k -le [Math]::Min(12, $r0 + $nr); $k++) { $rh += [Math]::Round([double]$ws.Rows.Item($k).RowHeight, 3) }
                    $res.sheets += [ordered]@{
                        name = $ws.Name; visible = [int]$ws.Visible; used = $ur.Address($false, $false); cells = $cells; colWidths = $cw; rowHeights = $rh
                        shapes = [int]$ws.Shapes.Count; charts = [int]$ws.ChartObjects().Count; tables = [int]$ws.ListObjects.Count; comments = [int]$ws.Comments.Count
                        hyperlinks = [int]$ws.Hyperlinks.Count; filter = [bool]$ws.AutoFilterMode; protectedSheet = [bool]$ws.ProtectContents
                        cf = [int]$ws.Cells.FormatConditions.Count
                    }
                }
                try { $win = $wb.Windows.Item(1); $res.freeze = [ordered]@{ sheet = $wb.ActiveSheet.Name; on = [bool]$win.FreezePanes; row = [int]$win.SplitRow; col = [int]$win.SplitColumn } } catch { $res.freeze = $null }
                $res.ok = (-not $res.repaired) -and (-not $res.error)
            }
        } catch { $res.error = $_.Exception.Message }
        finally { if ($wb) { try { $wb.Close($false) } catch {} } }
        $res | ConvertTo-Json -Depth 8 -Compress | Set-Content -Path (Join-Path $OutDir ($f.Name + '.json')) -Encoding UTF8
        Write-Host ("{0} ok={1} repaired={2}" -f $f.Name, $res.ok, $res.repaired)
    }
} finally { Stop-Xl $h }
