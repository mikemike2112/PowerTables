# Avvia l'app PowerTables su un file di lavoro, cattura SOLO la finestra dell'app e chiude il processo avviato (per PID).
param([string]$File = '', [Parameter(Mandatory)][string]$Out, [int]$WaitSec = 5, [string[]]$AppArgs = @(), [string]$Exe = '')
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Runtime.InteropServices;
public class PtWin {
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint flags);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
}
"@
if (-not $Exe) { $Exe = Join-Path $PSScriptRoot '..\src\PowerTables.App\bin\Debug\net10.0-windows\PowerTables.App.exe' }
$al = @(); if ($File) { $al += (Resolve-Path $File).Path }; $al += $AppArgs
$p = Start-Process -FilePath $Exe -ArgumentList $al -PassThru
try {
    Start-Sleep -Seconds $WaitSec
    $p.Refresh()
    $h = $p.MainWindowHandle
    if ($h -eq 0) { throw "finestra non trovata" }
    $r = New-Object PtWin+RECT; [void][PtWin]::GetWindowRect($h, [ref]$r)
    $w = $r.R - $r.L; $ht = $r.B - $r.T
    $bmp = New-Object System.Drawing.Bitmap $w, $ht
    $g = [System.Drawing.Graphics]::FromImage($bmp); $hdc = $g.GetHdc()
    [void][PtWin]::PrintWindow($h, $hdc, 2); $g.ReleaseHdc($hdc); $g.Dispose()
    $bmp.Save($Out, [System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
} finally {
    if (-not $p.HasExited) { $p.CloseMainWindow() | Out-Null; Start-Sleep -Milliseconds 800 }
    if (-not $p.HasExited) { Stop-Process -Id $p.Id -Force }
}
