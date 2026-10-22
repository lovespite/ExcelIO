using System.Globalization;
using System.IO.Compression;
using System.Text;
using ExcelIO;
using QuickExcel.Core;

[assembly: CollectionBehavior(DisableTestParallelization = true)]

namespace QuickExcel.Core.Test;

public sealed class WorkbookDocumentTests : IDisposable
{
    private readonly string _folder = Path.Combine(Path.GetTempPath(), "QuickExcel-" + Guid.NewGuid().ToString("N"));

    public WorkbookDocumentTests() => Directory.CreateDirectory(_folder);
    private string FilePath(string name) => Path.Combine(_folder, name);

    [Fact]
    public void BlankDocumentSupportsLiteralsAndFormulaEditing()
    {
        using var document = WorkbookDocument.CreateBlank();
        Assert.Single(document.Workbook);
        document.SetCell(0, 1, 0, "10");
        document.SetCell(0, 0, 0, "=A2*3");
        document.Calculate();
        Assert.Equal("=A2*3", document.GetCellInput(0, 0, 0));
        Assert.Equal("30", document.GetCellValue(0, 0, 0));
        document.SetCell(0, 0, 0, "'=A2*3");
        Assert.Equal("'=A2*3", document.GetCellInput(0, 0, 0));
        Assert.Equal("=A2*3", document.GetCellValue(0, 0, 0));
        Assert.False(document.Workbook[0][0].Cells[0].HasFormula);
    }

    [Fact]
    public void CalculationsFollowDependenciesAndBoundFullColumnRanges()
    {
        using var document = WorkbookDocument.CreateBlank();
        document.SetCell(0, 0, 0, "=B1+B1");
        document.SetCell(0, 0, 1, "=C1*2");
        document.SetCell(0, 0, 2, "4");
        document.SetCell(0, 1, 3, "=SUM(A)");
        document.Calculate();
        Assert.Equal("16", document.GetCellValue(0, 0, 0));
        Assert.Equal("16", document.GetCellValue(0, 1, 3));
        Assert.Empty(document.ScriptErrors);
    }

    [Fact]
    public void MalformedFormulaDoesNotPreventOtherCellsCalculating()
    {
        using var document = WorkbookDocument.CreateBlank();
        document.SetCell(0, 0, 0, "=$");
        document.SetCell(0, 0, 1, "=2+3");
        document.Calculate();
        Assert.Equal("#VALUE!", document.GetCellValue(0, 0, 0));
        Assert.Equal("5", document.GetCellValue(0, 0, 1));
    }

    [Fact]
    public void CircularDependenciesAreVisibleInsteadOfUsingStaleNumbers()
    {
        using var document = WorkbookDocument.CreateBlank();
        document.SetCell(0, 0, 0, "=B1");
        document.SetCell(0, 0, 1, "=A1");
        document.SetCell(0, 0, 2, "=A1+1");
        document.Calculate();
        Assert.Equal("#CYCLE!", document.GetCellValue(0, 0, 0));
        Assert.Equal("#CYCLE!", document.GetCellValue(0, 0, 2));
        Assert.NotEmpty(document.ScriptErrors);
    }

    [Fact]
    public void UnsupportedCrossSheetReferenceRetainsItsSavedResult()
    {
        using var document = WorkbookDocument.CreateBlank();
        document.SetCell(0, 0, 0, "=Sheet2!A1");
        document.Workbook[0][0].Cells[0].SetCalculatedValue("42");
        document.Calculate();
        Assert.Equal("42", document.GetCellValue(0, 0, 0));
        Assert.Contains(document.ScriptErrors, message => message.Contains("跨工作表"));
    }

    [Fact]
    public void FormulaNumbersUseInvariantCultureAndRestoreTheCallerCulture()
    {
        var previous = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("de-DE");
            using var document = WorkbookDocument.CreateBlank();
            document.SetCell(0, 0, 0, "1.5");
            document.SetCell(0, 0, 1, "=A1+0.25");
            document.Calculate();
            Assert.Equal("1.75", document.GetCellValue(0, 0, 1));
            Assert.Equal("de-DE", CultureInfo.CurrentCulture.Name);
        }
        finally { CultureInfo.CurrentCulture = previous; }
    }

    [Fact]
    public void DemoKeepsUsefulCachedValuesUntilScriptsAreEnabled()
    {
        using var document = WorkbookDocument.CreateDemo();
        Assert.False(document.ScriptsEnabled);
        Assert.Equal("1066.5", document.GetCellValue(0, 4, 4));
        document.SetCell(0, 1, 1, "100");
        document.Calculate();
        Assert.Equal("1066.5", document.GetCellValue(0, 4, 4));
        Assert.NotEmpty(document.ScriptErrors);
        document.EnableScriptsAndCalculate();
        Assert.True(document.ScriptsEnabled);
        Assert.Equal("180", document.GetCellValue(0, 1, 4));
        Assert.Equal("708.3", document.GetCellValue(0, 4, 4));
    }

    [Fact]
    public void JavascriptHandlesScalarAndRangeArgumentsAndRecompiles()
    {
        using var document = WorkbookDocument.CreateBlank();
        document.ScriptSource = "quickExcel.functions.register('DOUBLE', x => x * 2); quickExcel.functions.register('TOTAL', xs => xs.reduce((a,b) => a + Number(b), 0));";
        document.SetCell(0, 0, 0, "3");
        document.SetCell(0, 1, 0, "4");
        document.SetCell(0, 0, 1, "=DOUBLE(A1)");
        document.SetCell(0, 1, 1, "=TOTAL(A1:A2)");
        document.EnableScriptsAndCalculate();
        Assert.Equal("6", document.GetCellValue(0, 0, 1));
        Assert.Equal("7", document.GetCellValue(0, 1, 1));
        document.ScriptSource = "quickExcel.functions.register('DOUBLE', x => x * 3);";
        Assert.False(document.ScriptsEnabled);
        document.EnableScriptsAndCalculate();
        Assert.Equal("9", document.GetCellValue(0, 0, 1));
        Assert.Equal("#NAME?", document.GetCellValue(0, 1, 1));
    }

    [Fact]
    public void JavascriptErrorsProduceCellErrorAndDiagnostic()
    {
        using var document = WorkbookDocument.CreateBlank();
        document.ScriptSource = "quickExcel.functions.register('FAIL', () => { throw new Error('demo failure'); });";
        document.SetCell(0, 0, 0, "=FAIL()");
        document.EnableScriptsAndCalculate();
        Assert.Equal("#SCRIPT!", document.GetCellValue(0, 0, 0));
        Assert.Contains(document.ScriptErrors, message => message.Contains("demo failure"));
    }

    [Fact]
    public void CompilationFailureDisablesScriptsAndPreservesCachedResults()
    {
        using var document = WorkbookDocument.CreateDemo();
        document.ScriptSource = "throw new Error('compile failure');";
        Assert.ThrowsAny<Exception>(document.EnableScriptsAndCalculate);
        Assert.False(document.ScriptsEnabled);
        Assert.Equal("1066.5", document.GetCellValue(0, 4, 4));
        Assert.Contains(document.ScriptErrors, message => message.Contains("compile failure"));
    }

    [Fact]
    public void PackageRoundTripPreservesSourceFormulaCacheStyleAndSheetsWithoutExecutingScript()
    {
        var path = FilePath("roundtrip.xlsq");
        using (var original = WorkbookDocument.CreateDemo())
        {
            original.ScriptSource = "throw new Error('must never run during open');";
            original.AddSheet("第二张表");
            original.SetCell(1, 0, 0, "测试内容");
            original.Save(path);
        }
        using var reopened = WorkbookDocument.Open(path);
        Assert.False(reopened.ScriptsEnabled);
        Assert.Equal("throw new Error('must never run during open');", reopened.ScriptSource);
        Assert.Equal("=SUM(E2:E4)", reopened.GetCellInput(0, 4, 4));
        Assert.Equal("1066.5", reopened.GetCellValue(0, 4, 4));
        Assert.True(reopened.Workbook[0][0].Cells[0].Style?.Bold);
        Assert.Equal("测试内容", reopened.GetCellValue(1, 0, 0));
        Assert.Empty(reopened.ScriptErrors);
        using var archive = ZipFile.OpenRead(path);
        Assert.Equal(3, archive.Entries.Count);
        Assert.NotNull(archive.GetEntry("scripts/functions.js"));
    }

    [Fact]
    public void ExportWritesOnlyValuesWithoutChangingDocumentFormulas()
    {
        var path = FilePath("export.xlsx");
        using var document = WorkbookDocument.CreateDemo();
        document.EnableScriptsAndCalculate();
        document.ExportXlsx(path);
        using var imported = WorkbookDocument.Open(path);
        Assert.Equal("1066.5", imported.GetCellValue(0, 4, 4));
        Assert.DoesNotContain(imported.Workbook.SelectMany(sheet => sheet.Rows).SelectMany(row => row.Cells), cell => cell.HasFormula);
        Assert.Equal("=SUM(E2:E4)", document.GetCellInput(0, 4, 4));
        Assert.Equal("=DISCOUNT(B2,D2)*C2", document.GetCellInput(0, 1, 4));
    }

    [Fact]
    public void SparseRowsRetainOriginalColumnCoordinatesThroughBothFormats()
    {
        using var document = WorkbookDocument.CreateBlank();
        document.SetCell(0, 0, 4, "far right");
        document.SetCell(0, 3, 1, "=6*7");
        document.Calculate();
        var packagePath = FilePath("sparse.xlsq");
        document.Save(packagePath);
        using var reopened = WorkbookDocument.Open(packagePath);
        Assert.Equal("", reopened.GetCellValue(0, 0, 0));
        Assert.Equal("far right", reopened.GetCellValue(0, 0, 4));
        Assert.Equal("=6*7", reopened.GetCellInput(0, 3, 1));
        var exportPath = FilePath("sparse.xlsx");
        reopened.ExportXlsx(exportPath);
        using var exported = WorkbookDocument.Open(exportPath);
        Assert.Equal("42", exported.GetCellValue(0, 3, 1));
        Assert.Equal("", exported.GetCellValue(0, 3, 0));
    }

    [Fact]
    public void ImportedExcelFormulaAlwaysDisplaysLeadingEqualsInEditor()
    {
        var path = FilePath("formula.xlsx");
        var workbook = new XlWorkbook();
        var sheet = workbook.NewWorksheet("Sheet1");
        sheet.AddRow("42");
        sheet[0].Cells[0].SetFormula("6*7", "42");
        XlHelper.Save(path, workbook);
        using var document = WorkbookDocument.Open(path);
        Assert.Equal("=6*7", document.GetCellInput(0, 0, 0));
        document.SetCell(0, 0, 0, document.GetCellInput(0, 0, 0));
        document.Calculate();
        Assert.Equal("42", document.GetCellValue(0, 0, 0));
    }

    [Fact]
    public void SavingExistingPackageReplacesItAndFailedSaveKeepsPreviousFile()
    {
        var path = FilePath("safe.xlsq");
        using var document = WorkbookDocument.CreateBlank();
        document.SetCell(0, 0, 0, "first");
        document.Save(path);
        document.SetCell(0, 0, 0, "second");
        document.Save(path);
        document.ScriptSource = new string('x', 1024 * 1024 + 1);
        Assert.Throws<InvalidOperationException>(() => document.Save(path));
        using var reopened = WorkbookDocument.Open(path);
        Assert.Equal("second", reopened.GetCellValue(0, 0, 0));
        Assert.Empty(Directory.GetFiles(_folder, "*.tmp"));
    }

    [Theory]
    [InlineData("{\"formatVersion\":2}", "scripts/functions.js")]
    [InlineData("{\"formatVersion\":1}", "missing.js")]
    public void InvalidVersionOrMissingScriptFailsClearly(string manifest, string scriptEntry)
    {
        var path = FilePath("invalid.xlsq");
        using (var archive = ZipFile.Open(path, ZipArchiveMode.Create))
        {
            WriteText(archive, "manifest.json", manifest);
            WriteText(archive, scriptEntry, "");
        }
        Assert.Throws<InvalidDataException>(() => WorkbookDocument.Open(path));
    }

    [Fact]
    public void OversizeScriptAndDuplicateEntriesAreRejected()
    {
        var path = FilePath("duplicate.xlsq");
        using (var archive = ZipFile.Open(path, ZipArchiveMode.Create))
        {
            WriteText(archive, "manifest.json", "{\"formatVersion\":1}");
            WriteText(archive, "manifest.json", "{\"formatVersion\":1}");
        }
        Assert.Throws<InvalidDataException>(() => WorkbookDocument.Open(path));
        var hugePath = FilePath("oversize.xlsq");
        using (var archive = ZipFile.Open(hugePath, ZipArchiveMode.Create))
        {
            WriteText(archive, "manifest.json", "{\"formatVersion\":1}");
            WriteText(archive, "scripts/functions.js", new string('a', 1024 * 1024 + 1));
        }
        Assert.Throws<InvalidDataException>(() => WorkbookDocument.Open(hugePath));
    }

    [Fact]
    public void ImportRejectsRowsThatWouldForceHugeDenseAllocations()
    {
        var path = FilePath("huge.xlsx");
        using (var archive = ZipFile.Open(path, ZipArchiveMode.Create))
            WriteText(archive, "xl/worksheets/sheet1.xml", "<worksheet><sheetData><row r=\"1048576\"><c r=\"A1048576\"/></row></sheetData></worksheet>");
        Assert.Throws<InvalidDataException>(() => WorkbookDocument.Open(path));
    }

    [Fact]
    public void EmptyImportedWorkbookGetsOneEditableSheet()
    {
        var path = FilePath("empty.xlsx");
        XlHelper.Save(path, new XlWorkbook());
        using var document = WorkbookDocument.Open(path);
        Assert.Single(document.Workbook);
        document.SetCell(0, 0, 0, "ready");
        Assert.Equal("ready", document.GetCellValue(0, 0, 0));
    }

    [Fact]
    public void DocumentsDoNotAccumulateGlobalCellChangeHandlers()
    {
        var original = XlCell.OnValueChanged;
        Action<XlCell> external = _ => { };
        XlCell.OnValueChanged += external;
        try
        {
            var expected = XlCell.OnValueChanged;
            for (var index = 0; index < 20; index++)
            {
                using var document = WorkbookDocument.CreateBlank();
                document.DisableScripts();
            }
            Assert.Equal(expected, XlCell.OnValueChanged);
        }
        finally { XlCell.OnValueChanged = original; }
    }

    private static void WriteText(ZipArchive archive, string name, string text)
    {
        using var writer = new StreamWriter(archive.CreateEntry(name).Open(), new UTF8Encoding(false));
        writer.Write(text);
    }

    public void Dispose() => Directory.Delete(_folder, recursive: true);
}
