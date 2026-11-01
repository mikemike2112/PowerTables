using PowerTables.Core;
using PowerTables.Core.Formulas;

namespace PowerTables.Tests;

public class ItalianFormulaTests
{
    [Theory]
    [InlineData("SOMMA(A1;A2)", "SUM(A1,A2)")]
    [InlineData("SE(A1>1;\"sì\";\"no\")", "IF(A1>1,\"sì\",\"no\")")]
    [InlineData("SOMMA(A1;2,5)", "SUM(A1,2.5)")]
    [InlineData("SOMMA(A1:A3)", "SUM(A1:A3)")]
    [InlineData("CERCA.VERT(2;A1:B3;2;FALSO)", "VLOOKUP(2,A1:B3,2,FALSE)")]
    [InlineData("SE.ERRORE(1/0;VERO)", "IFERROR(1/0,TRUE)")]
    [InlineData("SUM(A1,A2)", "SUM(A1,A2)")]                       // inglese con virgola: invariato
    [InlineData("SUM({1,2;3,4})", "SUM({1,2;3,4})")]                // ';' dentro una costante matrice resta riga
    [InlineData("TESTO(1234,5;\"0\")", "TEXT(1234.5,\"0\")")]
    public void FromUser_TranslatesToNeutralForm(string input, string expected) => Assert.Equal(expected, FormulaText.FromUser(input));

    [Fact]
    public void TypedItalianFormula_CalculatesAndIsStoredInEnglish()
    {
        var wb = new Workbook(); var s = wb.AddSheet("S");
        wb.Do("x", () => { wb.SetValue(s, 1, 1, 4.0); wb.SetValue(s, 2, 1, 6.0); wb.SetFormula(s, 3, 1, "SOMMA(A1;A2;0,5)"); wb.SetFormula(s, 4, 1, "SE(A3>10;\"alto\";\"basso\")"); wb.SetFormula(s, 5, 1, "XLOOKUP(4;A1:A2;A1:A2)"); });
        Assert.Equal(10.5, s.GetValue(3, 1));
        Assert.Equal("alto", s.GetValue(4, 1));
        Assert.Equal("SUM(A1,A2,0.5)", s.Get(3, 1).F!.Text);          // nel file: nomi inglesi e virgola
        Assert.Equal("_xlfn.XLOOKUP(4,A1:A2,A1:A2)", s.Get(5, 1).F!.Text);
        Assert.Equal("=SUM(A1,A2,0.5)", wb.EditText(s, 3, 1));
    }

    [Fact]
    public void Catalog_CoversAllRegisteredFunctions_WithCategory()
    {
        Assert.Equal(Fns.Names.Distinct().Count(), FunctionCatalog.All.Count);
        foreach (var c in new[] { "Finanziarie", "Logiche", "Testo", "Data e ora" }) Assert.Contains(FunctionCatalog.All, f => f.Category == c);
        Assert.DoesNotContain(FunctionCatalog.All, f => f.Category == "Altre" && FunctionCatalog.EnglishOf(f.En) != null);
        Assert.Equal("SUM", FunctionCatalog.EnglishOf("somma"));
    }
}
