# Apre un .xlsx con Excel vero (istanza privata, nascosta), rileva errori/ripristino e scarica celle in JSON.
# Letture con Formula/Value2 (invarianti): mai FormulaLocal/Text.
param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Out, [switch]$Recalc, [int]$Timeout = 120)
. "$PSScriptRoot\Excel-Com.ps1"
$h = Start-Xl
$res = [ordered]@{ ok = $false; repaired = $false; error = $null; name = $null; sheets = @() }
try {
    $xl = $h.App
    $wb = $null
    try { $wb = $xl.Workbooks.Open((Resolve-Path $Path).Path, 0, $true) } catch { $res.error = $_.Exception.Message }
    if ($wb) {
        $res.name = $wb.Name
        $res.repaired = [bool]($wb.Name -match '\[(Repaired|Ripristinat\w*)\]')
        if ($Recalc) { $xl.CalculateFull() }
        foreach ($ws in $wb.Worksheets) {
            $ur = $ws.UsedRange
            $r0 = $ur.Row; $c0 = $ur.Column; $nr = $ur.Rows.Count; $nc = $ur.Columns.Count
            $cells = New-Object System.Collections.ArrayList
            $v2 = $ur.Value2; $fm = $ur.Formula; $nf = $null
            for ($i = 1; $i -le $nr; $i++) { for ($j = 1; $j -le $nc; $j++) {
                if ($nr -eq 1 -and $nc -eq 1) { $v = $v2; $f = $fm } else { $v = $v2[$i, $j]; $f = $fm[$i, $j] }
                if ($null -eq $v -and ($f -eq '' -or $null -eq $f)) { continue }
                $e = $null; $t = 'n'
                if ($v -is [int]) { $t = 'e'; $e = $v; $v = $null }
                elseif ($v -is [string]) { $t = 's' } elseif ($v -is [bool]) { $t = 'b' }
                $isF = ($f -is [string]) -and $f.StartsWith('=')
                [void]$cells.Add([ordered]@{ r = $r0 + $i - 1; c = $c0 + $j - 1; t = $t; v = $v; e = $e; f = $(if ($isF) { $f } else { $null }) })
            } }
            $merges = @(); $fr = 0; $fc = 0
            $res.sheets += [ordered]@{ name = $ws.Name; cells = $cells; visible = $ws.Visible }
        }
        $res.ok = (-not $res.repaired) -and (-not $res.error)
    }
} catch { $res.error = $_.Exception.Message }
finally { Stop-Xl $h }
$res | ConvertTo-Json -Depth 8 -Compress | Set-Content -Path $Out -Encoding UTF8
if (-not $res.ok) { Write-Error ("Excel: " + $res.error + " repaired=" + $res.repaired); exit 1 }
