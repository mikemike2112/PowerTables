# Helper COM per Excel (dot-source). Avvia un'istanza PRIVATA, la traccia per PID e la chiude sempre.
# Regole: mai Stop-Process -Name; kill solo del PID avviato qui. Finestra nascosta, alert disattivati.
# Le chiamate COM da PowerShell usano la cultura del thread: forzo en-US così NumberFormat/Formula sono invarianti (non toccano le impostazioni di Office).
[System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo('en-US')

function Start-Xl([bool]$Visible = $false) {
    $before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $Visible
    $xl.DisplayAlerts = $false
    $xl.AskToUpdateLinks = $false
    $xl.EnableEvents = $false
    $xl.AutomationSecurity = 3   # msoAutomationSecurityForceDisable: nessuna macro
    $pid_ = 0
    try {
        Add-Type -Namespace PT -Name Win -MemberDefinition '[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);' -ErrorAction SilentlyContinue
        [uint32]$p = 0; [void][PT.Win]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$p); $pid_ = [int]$p
    } catch {}
    if ($pid_ -eq 0) {
        $after = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
        $new = @($after | Where-Object { $before -notcontains $_ }); if ($new.Count -eq 1) { $pid_ = $new[0] }
    }
    if ($env:PT_PIDFILE -and $pid_ -gt 0) { Set-Content -Path $env:PT_PIDFILE -Value $pid_ }   # permette al chiamante di terminare SOLO questo PID in caso di blocco
    return [pscustomobject]@{ App = $xl; Pid = $pid_ }
}

function Stop-Xl($h) {
    if ($null -eq $h) { return }
    try { foreach ($w in @($h.App.Workbooks)) { try { $w.Close($false) } catch {} } } catch {}
    try { $h.App.Quit() } catch {}
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($h.App) } catch {}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($h.Pid -gt 0) {
        Start-Sleep -Milliseconds 700
        $p = Get-Process -Id $h.Pid -ErrorAction SilentlyContinue
        if ($p -and $p.ProcessName -eq 'EXCEL') { Stop-Process -Id $h.Pid -Force }   # solo il PID avviato da noi
    }
}

# Esegue un blocco con timeout/watchdog: se scade, uccide SOLO il PID dell'istanza.
function Invoke-XlGuarded($h, [scriptblock]$Block, [int]$TimeoutSec = 120) {
    $job = [powershell]::Create().AddScript($Block).AddArgument($h.App)
    $ar = $job.BeginInvoke()
    if (-not $ar.AsyncWaitHandle.WaitOne($TimeoutSec * 1000)) {
        if ($h.Pid -gt 0) { Stop-Process -Id $h.Pid -Force -ErrorAction SilentlyContinue }
        throw "Timeout COM dopo $TimeoutSec s (istanza $($h.Pid) terminata)"
    }
    $r = $job.EndInvoke($ar); $job.Dispose(); return $r
}
