# Genera tests/corpus/office-made/rich-text.xlsx (testo ricco in stringhe condivise) con Excel vero. Author/LastAuthor = "Corpus".
param([string]$OutDir = "$PSScriptRoot\..\tests\corpus\office-made")
. "$PSScriptRoot\Excel-Com.ps1"
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
$OutDir = (New-Item -ItemType Directory -Force $OutDir).FullName
$h = Start-Xl
try {
    $xl = $h.App; $wb = $xl.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
    $ws.Range('A1').Value2 = 'Grassetto e rosso corsivo'; $ws.Range('B1').Value2 = 42; $ws.Range('C1').Value2 = 'altra cella'
    $ws.Range('A1').Characters(1, 9).Font.Bold = $true
    $ws.Range('A1').Characters(13, 12).Font.Color = 255; $ws.Range('A1').Characters(13, 12).Font.Italic = $true
    $ws.Range('A2').Value2 = 'Solo la fine in grassetto'; $ws.Range('A2').Characters(15, 11).Font.Bold = $true; $ws.Range('B2').Value2 = 7
    $path = Join-Path $OutDir 'rich-text.xlsx'
    if (Test-Path $path) { Remove-Item $path -Force }
    $wb.SaveAs($path, 51); $wb.Close($false)
    $zip = [System.IO.Compression.ZipFile]::Open($path, 'Update')
    try {
        foreach ($n in 'docProps/core.xml') {
            $e = $zip.GetEntry($n); $sr = New-Object System.IO.StreamReader($e.Open()); $x = $sr.ReadToEnd(); $sr.Close()
            $x = [regex]::Replace($x, '<dc:creator>.*?</dc:creator>', '<dc:creator>Corpus</dc:creator>')
            $x = [regex]::Replace($x, '<cp:lastModifiedBy>.*?</cp:lastModifiedBy>', '<cp:lastModifiedBy>Corpus</cp:lastModifiedBy>')
            $e.Delete(); $ne = $zip.CreateEntry($n); $sw = New-Object System.IO.StreamWriter($ne.Open(), (New-Object System.Text.UTF8Encoding($false))); $sw.Write($x); $sw.Close()
        }
        $a = $zip.GetEntry('docProps/app.xml'); $sr = New-Object System.IO.StreamReader($a.Open()); $x = $sr.ReadToEnd(); $sr.Close()
        $x = [regex]::Replace($x, '<Company>.*?</Company>', '<Company></Company>'); $a.Delete()
        $na = $zip.CreateEntry('docProps/app.xml'); $sw = New-Object System.IO.StreamWriter($na.Open(), (New-Object System.Text.UTF8Encoding($false))); $sw.Write($x); $sw.Close()
    } finally { $zip.Dispose() }
    Write-Host "ok rich-text.xlsx"
} finally { Stop-Xl $h }
