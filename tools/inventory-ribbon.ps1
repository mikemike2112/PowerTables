# Inventario del ribbon di Excel via UI Automation (schede > gruppi > comandi). Excel visibile su un file vuoto creato qui.
param([string]$Out = "$PSScriptRoot\..\tests\work\ribbon-inventory.json")
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
. "$PSScriptRoot\Excel-Com.ps1"
$h = Start-Xl $true
$res = [ordered]@{ tabs = @() }
try {
    $xl = $h.App; $wb = $xl.Workbooks.Add(); $xl.WindowState = -4137
    Start-Sleep -Seconds 2
    $root = [System.Windows.Automation.AutomationElement]::FromHandle([IntPtr]$xl.Hwnd)
    $all = [System.Windows.Automation.TreeScope]::Descendants
    $ct = [System.Windows.Automation.ControlType]
    $cond = New-Object System.Windows.Automation.PropertyCondition ([System.Windows.Automation.AutomationElement]::ControlTypeProperty), $ct::TabItem
    $tabs = $root.FindAll($all, $cond)
    foreach ($tab in $tabs) {
        $name = $tab.Current.Name
        if (-not $name) { continue }
        try { $sel = $tab.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern); $sel.Select() } catch { continue }
        Start-Sleep -Milliseconds 700
        $root2 = [System.Windows.Automation.AutomationElement]::FromHandle([IntPtr]$xl.Hwnd)
        $groups = @()
        $gc = New-Object System.Windows.Automation.PropertyCondition ([System.Windows.Automation.AutomationElement]::ControlTypeProperty), $ct::Group
        foreach ($g in $root2.FindAll($all, $gc)) {
            $r = $g.Current.BoundingRectangle
            if ($r.Width -le 0 -or $g.Current.Name -eq '') { continue }
            $cmds = @()
            foreach ($t in @($ct::Button, $ct::SplitButton, $ct::CheckBox, $ct::MenuItem, $ct::ComboBox, $ct::RadioButton)) {
                $c = New-Object System.Windows.Automation.PropertyCondition ([System.Windows.Automation.AutomationElement]::ControlTypeProperty), $t
                foreach ($e in $g.FindAll($all, $c)) { $n = $e.Current.Name; if ($n -and $cmds -notcontains $n) { $cmds += $n } }
            }
            if ($cmds.Count -gt 0) { $groups += [ordered]@{ group = $g.Current.Name; commands = $cmds } }
        }
        $res.tabs += [ordered]@{ tab = $name; groups = $groups }
    }
    $wb.Close($false)
} catch { $res.error = $_.Exception.Message }
finally { Stop-Xl $h }
$res | ConvertTo-Json -Depth 6 | Set-Content -Path $Out -Encoding UTF8
