using PowerTables.Core;

namespace PowerTables.Tests;

public class DataToolsTests
{
    static (Workbook, Sheet) New() { var wb = new Workbook(); return (wb, wb.AddSheet("S")); }

    [Fact]
    public void PasteFormatsCopiesStyleOnlyAndUndoes()
    {
        var (wb, s) = New();
        wb.Do("x", () => { wb.SetValue(s, 1, 1, 1.0); wb.SetValue(s, 3, 3, "keep"); });
        wb.FormatRange(s, 1, 1, 1, 1, x => x with { Font = x.Font with { Bold = true } });
        wb.PasteFormats(s, 1, 1, 1, 1, s, 3, 3, 4, 4);
        Assert.True(wb.Styles.Get(s.Get(3, 3).Style).Font.Bold);
        Assert.True(wb.Styles.Get(s.Get(4, 4).Style).Font.Bold);
        Assert.Equal("keep", s.GetValue(3, 3));
        wb.Undo();
        Assert.False(wb.Styles.Get(s.Get(3, 3).Style).Font.Bold);
        Assert.Equal("keep", s.GetValue(3, 3));
    }

    [Fact]
    public void RemoveDuplicatesKeepsFirstAndShiftsUp()
    {
        var (wb, s) = New();
        wb.Do("x", () =>
        {
            string[] v = ["Nome", "a", "B", "A", "c", "b"];
            for (int i = 0; i < v.Length; i++) { wb.SetValue(s, i + 1, 1, v[i]); wb.SetValue(s, i + 1, 2, (double)i); }
            wb.SetFormula(s, 5, 3, "B5*2");
        });
        int n = wb.RemoveDuplicates(s, 1, 1, 6, 3, [1], true);
        Assert.Equal(2, n);
        Assert.Equal(["Nome", "a", "B", "c"], Enumerable.Range(1, 4).Select(r => s.GetValue(r, 1)).ToArray());
        Assert.Null(s.GetValue(5, 1)); Assert.Null(s.GetValue(6, 2));
        Assert.Equal(8.0, s.GetValue(4, 3));   // la formula "B5*2" è salita alla riga 4 e segue B4
        wb.Undo();
        Assert.Equal("A", s.GetValue(4, 1)); Assert.Equal("b", s.GetValue(6, 1)); Assert.Equal(8.0, s.GetValue(5, 3));
    }

    [Fact]
    public void RemoveDuplicatesNoKeyColumnsThrows()
    {
        var (wb, s) = New();
        Assert.Throws<ArgumentException>(() => wb.RemoveDuplicates(s, 1, 1, 3, 1, [], false));
    }

    [Fact]
    public void TextToColumnsSplitsAndTypesValues()
    {
        var (wb, s) = New();
        wb.Do("x", () => { wb.SetValue(s, 1, 1, "a;12;\"x;y\""); wb.SetValue(s, 2, 1, "b;;3,5"); });
        int n = wb.TextToColumns(s, 1, 1, 2, ";", ci: new System.Globalization.CultureInfo("it-IT"));
        Assert.Equal(3, n);
        Assert.Equal("a", s.GetValue(1, 1)); Assert.Equal(12.0, s.GetValue(1, 2)); Assert.Equal("x;y", s.GetValue(1, 3));
        Assert.Equal("b", s.GetValue(2, 1)); Assert.Null(s.GetValue(2, 2)); Assert.Equal(3.5, s.GetValue(2, 3));
        wb.Undo();
        Assert.Equal("a;12;\"x;y\"", s.GetValue(1, 1)); Assert.Null(s.GetValue(1, 2));
    }

    [Fact]
    public void SplitTextMergesConsecutive()
    {
        Assert.Equal(["a", "b"], Workbook.SplitText("a  b", " ", true, '"'));
        Assert.Equal(["a", "", "b"], Workbook.SplitText("a  b", " ", false, '"'));
    }

    [Fact]
    public void FormatCellsDialogChangeIsOneUndoStep()
    {
        var (wb, s) = New();
        wb.Do("x", () => wb.SetValue(s, 1, 1, 0.256));
        wb.FormatRange(s, 1, 1, 1, 1, x => x with
        {
            NumFmtId = 164, NumFmtCode = "0.00%", HAlign = "center", VAlign = "top", Wrap = true, Indent = 2, FillColor = "FFFFFF00",
            Font = x.Font with { Name = "Arial", Size = 14, Bold = true, Italic = true, Underline = true, Strike = true, Color = "FFFF0000" },
        });
        var xf = wb.Styles.Get(s.Get(1, 1).Style);
        Assert.Equal("0.00%", wb.Styles.FormatCode(s.Get(1, 1).Style)); Assert.Equal("25.60%".Replace('.', System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0]), wb.DisplayText(s, 1, 1));
        Assert.Equal(("center", "top", true, 2, "FFFFFF00"), (xf.HAlign, xf.VAlign, xf.Wrap, xf.Indent, xf.FillColor));
        Assert.Equal(new FontInfo("Arial", 14, true, true, true, true, "FFFF0000"), xf.Font);
        wb.Undo();
        Assert.Equal(0u, s.Get(1, 1).Style);
    }
}
