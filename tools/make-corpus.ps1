# Genera il corpus tests/corpus/office-made con Excel vero (COM). Rigenerabile. Istanza privata, nascosta, chiusa a fine lavoro.
# Metadati: Author/LastAuthor = "Corpus" (impostati nel core.xml del file generato, senza toccare le impostazioni di Office).
param([string]$OutDir = "$PSScriptRoot\..\tests\corpus\office-made")
. "$PSScriptRoot\Excel-Com.ps1"
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
$OutDir = (New-Item -ItemType Directory -Force $OutDir).FullName
$M = [Type]::Missing

function Set-CorpusMeta([string]$path) {
    $zip = [System.IO.Compression.ZipFile]::Open($path, 'Update')
    try {
        $e = $zip.GetEntry('docProps/core.xml')
        if ($e) {
            $sr = New-Object System.IO.StreamReader($e.Open()); $x = $sr.ReadToEnd(); $sr.Close()
            $x = [regex]::Replace($x, '<dc:creator>.*?</dc:creator>', '<dc:creator>Corpus</dc:creator>')
            $x = [regex]::Replace($x, '<cp:lastModifiedBy>.*?</cp:lastModifiedBy>', '<cp:lastModifiedBy>Corpus</cp:lastModifiedBy>')
            if ($x -notmatch 'lastModifiedBy') { $x = $x -replace '</cp:coreProperties>', '<cp:lastModifiedBy>Corpus</cp:lastModifiedBy></cp:coreProperties>' }
            $e.Delete()
            $ne = $zip.CreateEntry('docProps/core.xml'); $sw = New-Object System.IO.StreamWriter($ne.Open(), (New-Object System.Text.UTF8Encoding($false))); $sw.Write($x); $sw.Close()
        }
        Clear-CommentAuthors $zip
        $a = $zip.GetEntry('docProps/app.xml')
        if ($a) {
            $sr = New-Object System.IO.StreamReader($a.Open()); $x = $sr.ReadToEnd(); $sr.Close()
            $x = [regex]::Replace($x, '<Company>.*?</Company>', '<Company></Company>'); $x = [regex]::Replace($x, '<Manager>.*?</Manager>', '')
            $a.Delete()
            $na = $zip.CreateEntry('docProps/app.xml'); $sw = New-Object System.IO.StreamWriter($na.Open(), (New-Object System.Text.UTF8Encoding($false))); $sw.Write($x); $sw.Close()
        }
    } finally { $zip.Dispose() }
}

$h = Start-Xl
$count = 0
function New-Book($xl) { $wb = $xl.Workbooks.Add(); while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }; return $wb }
function Save-Book($xl, $wb, [string]$name) {
    $path = Join-Path $OutDir $name
    if (Test-Path $path) { Remove-Item $path -Force }
    $wb.SaveAs($path, 51)
    $wb.Close($false)
    Set-CorpusMeta $path
    $script:count++
    Write-Host "ok $name"
}
function Fill($ws, [string]$addr, $rows) {   # $rows: array di array
    $r0 = $ws.Range($addr).Row; $c0 = $ws.Range($addr).Column
    for ($i = 0; $i -lt $rows.Count; $i++) { for ($j = 0; $j -lt $rows[$i].Count; $j++) { $v = $rows[$i][$j]; if ($null -ne $v) { SetV $ws.Cells.Item($r0 + $i, $c0 + $j) ($v) } } }
}
$en = [System.Globalization.CultureInfo]::GetCultureInfo('en-US')
function SetNF($r, $f) { [void][System.__ComObject].InvokeMember('NumberFormat', [System.Reflection.BindingFlags]::SetProperty, $null, $r, @($f), $en) }
function SetV($cell, $v) { [void][System.__ComObject].InvokeMember('Value2', [System.Reflection.BindingFlags]::SetProperty, $null, $cell, @($v), $en) }
function F($ws, [string]$addr, [string]$f) { $ws.Range($addr).Formula = $f }

try {
    $xl = $h.App
    $data = @(@('Mela', 'x', 5, 44927), @('Banana', 'y', 3, 44958), @('Ciliegia', 'x', 8, 44989), @('Dattero', 'y', 1, 45020), @('Fico', 'x', 9, 45051), @('Uva', 'y', 2, 45082), @('Kiwi', 'x', 7, 45113), @('Pera', 'y', 4, 45144))

    # --- tipi e formati ---
    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    Fill $ws 'A1' @(@('Numero', 'Testo', 'Booleano', 'Data', 'Errore'), @(42, 'ciao', $true, 45000, $null), @(-3.14159, 'a "citato" & <tag>', $false, 1, $null), @(1E+20, "riga1`nriga2", $true, 60, $null), @(0.000001234, '', $true, 45000.75, $null))
    F $ws 'E2' '=1/0'; F $ws 'E3' '=NA()'; F $ws 'E4' '=#REF!'; SetNF $ws.Range('D2:D5') ('dd/mm/yyyy')
    Save-Book $xl $wb 'values-types.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    $fmts = @('General', '0', '0.00', '#,##0', '#,##0.00', '0%', '0.00%', '0.00E+00', '# ?/?', 'd/m/yyyy', 'd-mmm-yy', 'h:mm', 'h:mm:ss', 'h:mm AM/PM', '@', '$#,##0.00', '[$€-410] #,##0.00')
    for ($i = 0; $i -lt $fmts.Count; $i++) { SetV $ws.Cells.Item($i + 1, 1) (1234.5678); SetNF $ws.Cells.Item($i + 1, 1) ($fmts[$i]); SetV $ws.Cells.Item($i + 1, 2) ($fmts[$i]) }
    Save-Book $xl $wb 'numfmt-builtin.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    $cf = @('0.00;[Red]-0.00', '#,##0.0;(#,##0.0);"zero"', '0.0,"k"', '000000', '[<100]0.0;[>=100]0', 'dddd d mmmm yyyy', 'yyyy-mm-dd hh:mm', '"Tot: "0.00', '0.0%;-0.0%;;@', '[h]:mm:ss')
    for ($i = 0; $i -lt $cf.Count; $i++) { SetV $ws.Cells.Item($i + 1, 1) (12345.678); SetNF $ws.Cells.Item($i + 1, 1) ($cf[$i]); SetV $ws.Cells.Item($i + 1, 2) ($cf[$i]) }
    Save-Book $xl $wb 'numfmt-custom.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    Fill $ws 'A1' @(@(0), @(1), @(59), @(60), @(61), @(366), @(367), @(36526), @(44927), @(2958465)); SetNF $ws.Range('A1:A10') ('yyyy-mm-dd')
    F $ws 'B1' '=A1+0'; $ws.Range('B1').AutoFill($ws.Range('B1:B10')) | Out-Null
    F $ws 'C1' '=DATE(1900,2,29)'; F $ws 'C2' '=DAY(C1)'; F $ws 'C3' '=WEEKDAY(60)'; F $ws 'C4' '=DATE(1900,3,1)-DATE(1900,2,28)'
    Save-Book $xl $wb 'dates-1900.xlsx'

    $wb = New-Book $xl; $wb.Date1904 = $true; $ws = $wb.Worksheets.Item(1)
    Fill $ws 'A1' @(@(0), @(1), @(43465)); SetNF $ws.Range('A1:A3') ('yyyy-mm-dd'); F $ws 'B1' '=A3+1'
    Save-Book $xl $wb 'dates-1904.xlsx'

    # --- formule ---
    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    Fill $ws 'A1' @(@(10, 3), @(7, 2)); F $ws 'C1' '=A1+B1'; F $ws 'C2' '=A1-B1*2'; F $ws 'C3' '=A1/B1'; F $ws 'C4' '=A1^B1'; F $ws 'C5' '=-A1%'; F $ws 'C6' '=(A1+B1)*(A2-B2)'; F $ws 'C7' '=A1&"-"&B1'; F $ws 'C8' '=A1>B1'; F $ws 'C9' '=A1/0'
    Save-Book $xl $wb 'formula-arith.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    for ($i = 0; $i -lt $data.Count; $i++) { for ($j = 0; $j -lt 4; $j++) { SetV $ws.Cells.Item($i + 1, $j + 1) ($data[$i][$j]) } }
    F $ws 'F1' '=SUM(C1:C8)'; F $ws 'F2' '=AVERAGE(C1:C8)'; F $ws 'F3' '=MIN(C1:C8)'; F $ws 'F4' '=MAX(C1:C8)'; F $ws 'F5' '=COUNT(C1:D8)'; F $ws 'F6' '=COUNTA(A1:A8)'; F $ws 'F7' '=MEDIAN(C1:C8)'; F $ws 'F8' '=STDEV(C1:C8)'; F $ws 'F9' '=SUMPRODUCT(C1:C8,C1:C8)'; F $ws 'F10' '=LARGE(C1:C8,2)'; F $ws 'F11' '=SUBTOTAL(9,C1:C8)'
    Save-Book $xl $wb 'formula-sum-avg.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    Fill $ws 'A1' @(@(5), @(15), @(25)); F $ws 'B1' '=IF(A1>10,"alto","basso")'; $ws.Range('B1').AutoFill($ws.Range('B1:B3')) | Out-Null
    F $ws 'C1' '=AND(A1>0,A2>10)'; F $ws 'C2' '=OR(A1>100,A2>100)'; F $ws 'C3' '=NOT(A1>1)'; F $ws 'C4' '=IFERROR(1/0,"errore")'; F $ws 'C5' '=IFS(A1>20,"a",A1>2,"b")'; F $ws 'C6' '=SWITCH(A1,5,"cinque",15,"quindici","altro")'; F $ws 'C7' '=CHOOSE(2,"a","b","c")'; F $ws 'C8' '=XOR(TRUE,FALSE)'
    Save-Book $xl $wb 'formula-if-logic.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    Fill $ws 'A1' @(@('a', 100), @('b', 200), @('c', 300), @('d', 400)); F $ws 'D1' '=VLOOKUP("c",A1:B4,2,FALSE)'; F $ws 'D2' '=VLOOKUP("z",A1:B4,2,FALSE)'; F $ws 'D3' '=VLOOKUP("c",A1:B4,2)'; F $ws 'D4' '=HLOOKUP("b",A1:B2,2,FALSE)'; F $ws 'D5' '=LOOKUP("b",A1:A4,B1:B4)'
    Save-Book $xl $wb 'formula-vlookup.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    Fill $ws 'A1' @(@('a', 100), @('b', 200), @('c', 300)); F $ws 'D1' '=XLOOKUP("b",A1:A3,B1:B3)'; F $ws 'D2' '=XLOOKUP("z",A1:A3,B1:B3,"nd")'; F $ws 'D3' '=XMATCH("c",A1:A3)'; F $ws 'D4' '=IFNA(MATCH("z",A1:A3,0),-1)'
    Save-Book $xl $wb 'formula-xlookup.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    for ($i = 0; $i -lt $data.Count; $i++) { for ($j = 0; $j -lt 3; $j++) { SetV $ws.Cells.Item($i + 1, $j + 1) ($data[$i][$j]) } }
    F $ws 'E1' '=MATCH("Kiwi",A1:A8,0)'; F $ws 'E2' '=INDEX(C1:C8,E1)'; F $ws 'E3' '=INDEX(A1:C8,3,3)'; F $ws 'E4' '=INDEX(C1:C8,MATCH(MAX(C1:C8),C1:C8,0))'; F $ws 'E5' '=MATCH(4,C1:C8,1)'
    Save-Book $xl $wb 'formula-index-match.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    for ($i = 0; $i -lt $data.Count; $i++) { for ($j = 0; $j -lt 3; $j++) { SetV $ws.Cells.Item($i + 1, $j + 1) ($data[$i][$j]) } }
    F $ws 'E1' '=COUNTIF(B1:B8,"x")'; F $ws 'E2' '=SUMIF(B1:B8,"y",C1:C8)'; F $ws 'E3' '=COUNTIFS(B1:B8,"x",C1:C8,">5")'; F $ws 'E4' '=SUMIFS(C1:C8,B1:B8,"x",C1:C8,"<9")'; F $ws 'E5' '=AVERAGEIF(B1:B8,"x",C1:C8)'; F $ws 'E6' '=MAXIFS(C1:C8,B1:B8,"y")'; F $ws 'E7' '=COUNTIF(A1:A8,"*i*")'
    Save-Book $xl $wb 'formula-countif.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    $ws.Range('A1').Value2 = '  Hello World  '; $ws.Range('A2').Value2 = 'abc-def-ghi'
    F $ws 'B1' '=TRIM(A1)'; F $ws 'B2' '=LEN(A1)'; F $ws 'B3' '=UPPER(A2)'; F $ws 'B4' '=PROPER(A2)'; F $ws 'B5' '=LEFT(A2,3)&RIGHT(A2,3)'; F $ws 'B6' '=MID(A2,5,3)'; F $ws 'B7' '=FIND("-",A2)'; F $ws 'B8' '=SUBSTITUTE(A2,"-","/")'; F $ws 'B9' '=TEXTJOIN("|",TRUE,A1:A2)'; F $ws 'B10' '=CONCAT("a","b",1)'; F $ws 'B11' '=REPT("ab",3)'; F $ws 'B12' '=TEXT(1234,"0000")'; F $ws 'B13' '=VALUE("123")+1'; F $ws 'B14' '=EXACT("a","A")'
    Save-Book $xl $wb 'formula-text.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    $ws.Range('A1').Value2 = 45000; SetNF $ws.Range('A1') ('dd/mm/yyyy')
    F $ws 'B1' '=YEAR(A1)'; F $ws 'B2' '=MONTH(A1)'; F $ws 'B3' '=DAY(A1)'; F $ws 'B4' '=WEEKDAY(A1,2)'; F $ws 'B5' '=EDATE(A1,3)'; F $ws 'B6' '=EOMONTH(A1,1)'; F $ws 'B7' '=DATEDIF(A1,A1+400,"M")'; F $ws 'B8' '=NETWORKDAYS(A1,A1+30)'; F $ws 'B9' '=WORKDAY(A1,10)'; F $ws 'B10' '=DATE(2024,2,29)+1'; F $ws 'B11' '=WEEKNUM(A1)'; F $ws 'B12' '=TIME(13,30,0)'
    SetNF $ws.Range('B5:B6') ('dd/mm/yyyy'); SetNF $ws.Range('B10') ('dd/mm/yyyy'); SetNF $ws.Range('B12') ('hh:mm')
    Save-Book $xl $wb 'formula-date.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    F $ws 'A1' '=ROUND(2.675,2)'; F $ws 'A2' '=ROUNDUP(2.1,0)'; F $ws 'A3' '=MOD(-10,3)'; F $ws 'A4' '=POWER(2,10)'; F $ws 'A5' '=SQRT(2)'; F $ws 'A6' '=INT(-2.5)'; F $ws 'A7' '=ABS(-7)'; F $ws 'A8' '=CEILING(2.1,0.5)'; F $ws 'A9' '=FLOOR(2.9,0.5)'; F $ws 'A10' '=PI()'; F $ws 'A11' '=EXP(1)'; F $ws 'A12' '=LOG(1000,10)'; F $ws 'A13' '=GCD(12,18)'; F $ws 'A14' '=FACT(10)'; F $ws 'A15' '=PMT(0.05/12,60,10000)'
    Save-Book $xl $wb 'formula-math.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1); $ws.Name = 'Origine'; $ws2 = $wb.Worksheets.Add([Type]::Missing, $ws); $ws2.Name = 'Calcoli e dati'
    Fill $ws 'A1' @(@(1), @(2), @(3)); F $ws2 'A1' '=SUM(Origine!A1:A3)'; F $ws2 'A2' "=Origine!A1*10"; F $ws2 'A3' "=SUM('Calcoli e dati'!A1:A2)"
    Save-Book $xl $wb 'formula-cross-sheet.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    Fill $ws 'A1' @(@(2), @(4), @(6)); $ws.Range('D1').Value2 = 10; F $ws 'B1' '=A1*$D$1'; F $ws 'C1' '=$A1+B$1'; $ws.Range('B1:C1').AutoFill($ws.Range('B1:C3')) | Out-Null
    Save-Book $xl $wb 'formula-absolute-relative.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    for ($i = 1; $i -le 100; $i++) { SetV $ws.Cells.Item($i, 1) ($i); SetV $ws.Cells.Item($i, 2) ($i * 0.5) }
    F $ws 'C1' '=A1*B1+1'; $ws.Range('C1').AutoFill($ws.Range('C1:C100')) | Out-Null; F $ws 'D1' '=SUM($C$1:C1)'; $ws.Range('D1').AutoFill($ws.Range('D1:D100')) | Out-Null
    Save-Book $xl $wb 'formula-shared-fill.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    Fill $ws 'A1' @(@(1), @(2), @(3)); $wb.Names.Add('Valori', '=Foglio1!$A$1:$A$3') | Out-Null; $wb.Names.Add('Tasso', '=22/100') | Out-Null
    F $ws 'C1' '=SUM(Valori)'; F $ws 'C2' '=C1*(1+Tasso)'
    Save-Book $xl $wb 'formula-named-range.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    Fill $ws 'A1' @(@(1, 10), @(2, 20), @(3, 30)); $ws.Range('D1').FormulaArray = '=SUM(A1:A3*B1:B3)'; $ws.Range('D2:D4').FormulaArray = '=A1:A3*2'
    Save-Book $xl $wb 'formula-array-cse.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    Fill $ws 'A1' @(@(3), @(1), @(2)); $ws.Range('C1').Formula2 = '=SEQUENCE(5)'; $ws.Range('E1').Formula2 = '=SORT(A1:A3)'; $ws.Range('G1').Formula2 = '=UNIQUE(A1:A3)'; $ws.Range('I1').Formula = '=SUM(C1:C5)'
    Save-Book $xl $wb 'formula-dynamic-array.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    F $ws 'A1' '=1/0'; F $ws 'A2' '=NA()'; F $ws 'A3' '=VALUE("x")'; F $ws 'A4' '=SQRT(-1)'; F $ws 'A5' '=UNKNOWNFN(1)'; F $ws 'A6' '=#NULL!'; F $ws 'A7' '=SUM(A1,1)'; F $ws 'A8' '=IFERROR(A1,0)'; F $ws 'A9' '=ISERROR(A2)'
    Save-Book $xl $wb 'formula-errors.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    F $ws 'A1' '=TODAY()'; F $ws 'A2' '=NOW()'; F $ws 'A3' '=RAND()'; $ws.Range('B1').Value2 = 5; $ws.Range('B2').Value2 = 6; F $ws 'A4' '=OFFSET(B1,1,0)'; F $ws 'A5' '=INDIRECT("B1")+1'
    Save-Book $xl $wb 'formula-volatile.xlsx'

    # --- formattazione e layout ---
    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    $ws.Range('A1').Value2 = 'Grassetto'; $ws.Range('A1').Font.Bold = $true; $ws.Range('A2').Value2 = 'Corsivo rosso'; $ws.Range('A2').Font.Italic = $true; $ws.Range('A2').Font.Color = 255
    $ws.Range('A3').Value2 = 'Grande Arial'; $ws.Range('A3').Font.Name = 'Arial'; $ws.Range('A3').Font.Size = 18; $ws.Range('A4').Value2 = 'Sottolineato barrato'; $ws.Range('A4').Font.Underline = 2; $ws.Range('A4').Font.Strikethrough = $true
    $ws.Range('B1').Interior.Color = 65535; $ws.Range('B2').Interior.Color = 13421619; $ws.Range('B3').Interior.Color = 5287936; $ws.Range('B1:B3').Value2 = 'fill'
    Save-Book $xl $wb 'format-fonts-fills.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    $r = $ws.Range('B2:D4'); $r.Borders.LineStyle = 1; $ws.Range('B2:D2').Borders.Item(8).Weight = -4138; $ws.Range('F2').Borders.Item(9).LineStyle = -4115; $ws.Range('F2').Value2 = 'tratteggio'; $ws.Range('F4').Borders.Item(7).Weight = 4; $ws.Range('F4').Borders.Item(7).Color = 255; $ws.Range('F4').Value2 = 'rosso'
    Save-Book $xl $wb 'format-borders.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    $ws.Range('A1').Value2 = 'Testo molto lungo che va a capo nella cella stretta'; $ws.Range('A1').WrapText = $true; $ws.Columns.Item(1).ColumnWidth = 14
    $ws.Range('B1').Value2 = 'Dx'; $ws.Range('B1').HorizontalAlignment = -4152; $ws.Range('B2').Value2 = 'Centro'; $ws.Range('B2').HorizontalAlignment = -4108; $ws.Range('B3').Value2 = 'Alto'; $ws.Range('B3').VerticalAlignment = -4160; $ws.Rows.Item(3).RowHeight = 40; $ws.Range('B4').Value2 = 'Rientro'; $ws.Range('B4').IndentLevel = 2
    Save-Book $xl $wb 'format-alignment-wrap.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    $ws.Range('A1:C1').Merge(); $ws.Range('A1').Value2 = 'Titolo unito'; $ws.Range('A1').HorizontalAlignment = -4108; $ws.Range('A3:A5').Merge(); $ws.Range('A3').Value2 = 'verticale'; $ws.Range('C3:D4').Merge(); $ws.Range('C3').Value2 = 'blocco'
    Save-Book $xl $wb 'merged-cells.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    $ws.Range('A1:E3').Value2 = 'x'; $ws.Columns.Item(1).ColumnWidth = 5; $ws.Columns.Item(2).ColumnWidth = 30; $ws.Columns.Item(4).ColumnWidth = 0.5; $ws.Rows.Item(2).RowHeight = 45; $ws.Rows.Item(3).RowHeight = 8
    Save-Book $xl $wb 'col-widths-row-heights.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    for ($i = 1; $i -le 60; $i++) { for ($j = 1; $j -le 8; $j++) { SetV $ws.Cells.Item($i, $j) ($i * $j) } }
    $xl.Visible = $true; $ws.Activate(); $xl.ActiveWindow.FreezePanes = $false; $ws.Range('C3').Select(); $xl.ActiveWindow.FreezePanes = $true; $xl.Visible = $false
    Save-Book $xl $wb 'freeze-panes.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    $ws.Range('A1:C1').Value2 = [object[]]@('Nome', 'Gruppo', 'Valore')
    for ($i = 0; $i -lt $data.Count; $i++) { SetV $ws.Cells.Item($i + 2, 1) ($data[$i][0]); SetV $ws.Cells.Item($i + 2, 2) ($data[$i][1]); SetV $ws.Cells.Item($i + 2, 3) ($data[$i][2]) }
    $ws.Range('A1:C9').AutoFilter(2, 'x') | Out-Null
    Save-Book $xl $wb 'sort-filter.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1); $ws.Name = 'Uno'; $w2 = $wb.Worksheets.Add([Type]::Missing, $ws); $w2.Name = 'Due'; $w3 = $wb.Worksheets.Add([Type]::Missing, $w2); $w3.Name = 'Nascosto'
    $ws.Range('A1').Value2 = 'uno'; $w2.Range('A1').Value2 = 'due'; $w3.Range('A1').Value2 = 'nascosto'; $w3.Visible = 0; $ws.Tab.Color = 65280
    Save-Book $xl $wb 'multi-sheet.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    for ($i = 1; $i -le 10; $i++) { for ($j = 1; $j -le 5; $j++) { SetV $ws.Cells.Item($i, $j) ("$i-$j") } }
    $ws.Rows.Item(3).Hidden = $true; $ws.Columns.Item(2).Hidden = $true; $ws.Rows.Item('5:6').Hidden = $true
    Save-Book $xl $wb 'hidden-rows-cols.xlsx'

    # --- parti aggiuntive da preservare ---
    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    $ws.Range('A1:D1').Value2 = [object[]]@('Prodotto', 'Q1', 'Q2', 'Totale')
    for ($i = 0; $i -lt 5; $i++) { SetV $ws.Cells.Item($i + 2, 1) ("P$i"); SetV $ws.Cells.Item($i + 2, 2) (10 * ($i + 1)); SetV $ws.Cells.Item($i + 2, 3) (7 * ($i + 1)) }
    $lo = $ws.ListObjects.Add(1, $ws.Range('A1:D6'), $M, 1); $lo.Name = 'Vendite'; $lo.ListColumns.Item(4).DataBodyRange.Formula = '=[@Q1]+[@Q2]'
    Save-Book $xl $wb 'table-listobject.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    Fill $ws 'A1' @(@('Mese', 'Valore'), @('Gen', 10), @('Feb', 20), @('Mar', 15), @('Apr', 30))
    $co = $ws.Shapes.AddChart2(201, 51, 150, 10, 320, 200); $co.Chart.SetSourceData($ws.Range('A1:B5'))
    Save-Book $xl $wb 'chart-column.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    for ($i = 1; $i -le 10; $i++) { SetV $ws.Cells.Item($i, 1) ($i) }
    $fc = $ws.Range('A1:A10').FormatConditions.Add(1, 5, '=5'); $fc.Interior.Color = 255; $db = $ws.Range('B1:B10').FormatConditions.AddDatabar(); $ws.Range('B1:B10').Formula = '=A1*3'
    $cs = $ws.Range('C1:C10').FormatConditions.AddColorScale(2); $ws.Range('C1:C10').Formula = '=A1^2'
    Save-Book $xl $wb 'conditional-format.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    $ws.Range('A1:A3').Validation.Add(3, 1, 1, 'rosso,verde,blu') | Out-Null; $ws.Range('B1:B3').Validation.Add(1, 1, 1, '1', '10') | Out-Null; $ws.Range('A1').Value2 = 'rosso'; $ws.Range('B1').Value2 = 5
    Save-Book $xl $wb 'data-validation.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    $ws.Range('A1').Value2 = 'con nota'; $ws.Range('A1').AddComment('Questa è una nota di prova') | Out-Null
    Save-Book $xl $wb 'comments-notes.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    $ws.Hyperlinks.Add($ws.Range('A1'), 'https://example.org/', $M, 'Esempio', 'sito') | Out-Null; $ws.Hyperlinks.Add($ws.Range('A2'), '', "'Foglio1'!B5", $M, 'interno') | Out-Null
    Save-Book $xl $wb 'hyperlinks.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    $ws.Range('A1').Value2 = 'protetto'; $ws.Range('B1').Locked = $false; $ws.Protect()
    Save-Book $xl $wb 'sheet-protection.xlsx'

    # --- grandi e misti ---
    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1)
    $arr = New-Object 'object[,]' 1000, 4
    for ($i = 0; $i -lt 1000; $i++) { $arr[$i, 0] = $i + 1; $arr[$i, 1] = "Voce " + ($i % 37); $arr[$i, 2] = [math]::Round(($i * 7.31) % 1000, 2); $arr[$i, 3] = 44927 + ($i % 365) }
    $ws.Range('A2:D1001').Value2 = $arr; $ws.Range('A1:E1').Value2 = [object[]]@('ID', 'Voce', 'Importo', 'Data', 'IVA'); $ws.Range('E2:E1001').Formula = '=C2*0.22'; SetNF $ws.Range('D2:D1001') ('dd/mm/yyyy'); $ws.Range('G1').Formula = '=SUM(E2:E1001)'; $ws.Range('G2').Formula = '=COUNTIF(B2:B1001,"Voce 5")'
    Save-Book $xl $wb 'large-1000-rows.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1); $ws.Name = 'Vendite'
    $ws.Range('A1').Value2 = 'Rapporto vendite 2024'; $ws.Range('A1').Font.Size = 16; $ws.Range('A1').Font.Bold = $true; $ws.Range('A1:E1').Merge()
    $ws.Range('A3:E3').Value2 = [object[]]@('Regione', 'Prodotto', 'Unità', 'Prezzo', 'Ricavo'); $ws.Range('A3:E3').Font.Bold = $true; $ws.Range('A3:E3').Interior.Color = 14277081
    $reg = @('Nord', 'Sud', 'Centro'); $prd = @('A', 'B', 'C', 'D')
    for ($i = 0; $i -lt 12; $i++) { $r = $i + 4; SetV $ws.Cells.Item($r, 1) ($reg[$i % 3]); SetV $ws.Cells.Item($r, 2) ($prd[$i % 4]); SetV $ws.Cells.Item($r, 3) (10 + $i * 3); SetV $ws.Cells.Item($r, 4) (9.99 + $i); $ws.Cells.Item($r, 5).Formula = "=C$r*D$r" }
    SetNF $ws.Range('D4:E15') ('#,##0.00'); $ws.Range('A17').Value2 = 'Totale'; $ws.Range('E17').Formula = '=SUM(E4:E15)'; $ws.Range('E17').Font.Bold = $true; $ws.Range('E17').Borders.Item(8).LineStyle = 1
    $ws.Range('G3').Value2 = 'Per regione'; $ws.Range('G4').Value2 = 'Nord'; $ws.Range('G5').Value2 = 'Sud'; $ws.Range('G6').Value2 = 'Centro'; $ws.Range('H4:H6').Formula = '=SUMIF($A$4:$A$15,G4,$E$4:$E$15)'
    $co = $ws.Shapes.AddChart2(201, 51, 480, 40, 300, 180); $co.Chart.SetSourceData($ws.Range('G4:H6'))
    $ws.Columns.Item('A:E').ColumnWidth = 12; $xl.Visible = $true; $ws.Activate(); $ws.Range('A4').Select(); $xl.ActiveWindow.FreezePanes = $true; $xl.Visible = $false
    Save-Book $xl $wb 'mixed-sales-report.xlsx'

    $wb = New-Book $xl; $ws = $wb.Worksheets.Item(1); $ws.Name = 'Budget'; $w2 = $wb.Worksheets.Add([Type]::Missing, $ws); $w2.Name = 'Riepilogo'
    $ws.Range('A1:D1').Value2 = [object[]]@('Voce', 'Gen', 'Feb', 'Totale'); $ws.Range('A1:D1').Font.Bold = $true
    $v = @('Affitto', 'Utenze', 'Stipendi', 'Marketing')
    for ($i = 0; $i -lt 4; $i++) { SetV $ws.Cells.Item($i + 2, 1) ($v[$i]); SetV $ws.Cells.Item($i + 2, 2) (1000 * ($i + 1)); SetV $ws.Cells.Item($i + 2, 3) (1100 * ($i + 1)); $ws.Cells.Item($i + 2, 4).Formula = "=SUM(B$($i + 2):C$($i + 2))" }
    SetNF $ws.Range('B2:D5') ('#,##0 "€"'); $w2.Range('A1').Value2 = 'Totale annuo stimato'; $w2.Range('B1').Formula = '=SUM(Budget!D2:D5)*6'; SetNF $w2.Range('B1') ('#,##0 "€"'); $w2.Range('A2').Value2 = 'Voce maggiore'; $w2.Range('B2').Formula = '=INDEX(Budget!A2:A5,MATCH(MAX(Budget!D2:D5),Budget!D2:D5,0))'
    $ws.Range('A1:D5').Borders.LineStyle = 1; $names = $wb.Names.Add('TotaleAnnuo', '=Riepilogo!$B$1')
    Save-Book $xl $wb 'mixed-budget.xlsx'
} finally { Stop-Xl $h }
Write-Host "File generati: $count"
