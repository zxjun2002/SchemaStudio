using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SchemaStudio.Models;
using SchemaStudio.Services;
using SchemaStudio.UI;

internal static class Program
{
    private static readonly string Output = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../../artifacts/checks"));
    private static int passed;
    [STAThread]
    private static int Main(string[] args)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        Directory.CreateDirectory(Output);
        try
        {
            Check("CSV whitespace, text formulas, multiline and encoding round trip", CsvRoundTrip);
            Check("Column metadata survives rename, delete, undo and save", ColumnHistory);
            Check("XLSX data is read-only while schema metadata remains editable", XlsxReadOnlyAndSchema);
            Check("Formula protection and CSV literal equals text", FormulaProtection);
            Check("Source and schema conflicts cannot overwrite external changes", ConflictProtection);
            Check("Schema-only saves leave source workbook byte-identical", SchemaOnly);
            Check("Failed schema write restores source and keeps edit history", Rollback);
            Check("Legacy XLS is read-only, while metadata remains editable", Legacy);
            Check("Validation uses current values and detects composite primary keys", Validation);
            Check("Shared, array, external and merged formula regions stay protected", SpecialRanges);
            Check("Auxiliary and empty sheets cannot corrupt the editable table", MixedSheets);
            Check("Unconfigured columns auto-infer types on open; manual add requires type", AutoInferredTypes);
            Check("Opening workspace directory auto-generates and supplements missing schemas", WorkspaceSchemaSync);
            Check("History tracks up to 100 steps and supports direct jump to any step", HistoryJumpAndLimit);
            if (args.Contains("--ui")) Check("Desktop layout, virtual grid editing, formula lock and multi-sheet state", Ui);
            Console.WriteLine($"PASS: {passed} checks. Artifacts: {Output}"); return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }
    private static void Check(string name, Action action) { action(); passed++; Console.WriteLine("PASS " + name); }
    private static void Assert(bool value, string message) { if (!value) throw new Exception(message); }
    private static ColumnSchema Column(string name) => new() { ColName = name, ValueType = "string", ColProperties = new(), IsConfigured = true };
    private static void Throws(Action action, string message)
    {
        try { action(); } catch (InvalidOperationException) { return; } catch (IOException) { return; }
        throw new Exception(message);
    }
    private static string Csv(string name)
    {
        string path = Path.Combine(Output, name + "-" + Guid.NewGuid().ToString("N")[..6] + ".csv");
        File.WriteAllText(path, "id,name,price\r\n001,  木剑  ,100\r\n002,盾牌,200\r\n", new UTF8Encoding(true));
        return path;
    }
    private static string Xlsx(string name)
    {
        string path = Path.Combine(Output, name + "-" + Guid.NewGuid().ToString("N")[..6] + ".xlsx");
        using (var workbook = new XSSFWorkbook())
        {
            var sheet = workbook.CreateSheet("Items");
            var header = sheet.CreateRow(0);
            string[] names = { "id", "name", "price", "count", "total" };
            for (int i = 0; i < names.Length; i++) header.CreateCell(i).SetCellValue(names[i]);
            for (int i = 1; i <= 8; i++)
            {
                var row = sheet.CreateRow(i);
                row.CreateCell(0).SetCellValue(i); row.CreateCell(1).SetCellValue(i % 2 == 0 ? "守卫之盾" : "旅行者长剑");
                row.CreateCell(2).SetCellValue(i * 120); row.CreateCell(3).SetCellValue(i + 1); row.CreateCell(4).SetCellFormula($"C{i + 1}*D{i + 1}");
            }
            var style = workbook.CreateCellStyle(); style.DataFormat = workbook.CreateDataFormat().GetFormat("0.00");
            sheet.GetRow(1).GetCell(2).CellStyle = style;
            var lookup = workbook.CreateSheet("Lookup"); lookup.CreateRow(0).CreateCell(0).SetCellValue("value");
            lookup.CreateRow(1).CreateCell(0).SetCellFormula("Items!C2");
            workbook.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll();
            using var stream = File.Create(path); workbook.Write(stream);
        }
        // Include a real external-workbook relationship and an arbitrary extension part.
        using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            using (var writer = new StreamWriter(zip.CreateEntry("customXml/sentinel.xml").Open())) writer.Write("<preserve>do not rebuild this file</preserve>");
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XNamespace relationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XDocument Read(string name) { using var input = zip.GetEntry(name)!.Open(); return XDocument.Load(input); }
            void Write(string name, XDocument xml) { zip.GetEntry(name)?.Delete(); using var output = zip.CreateEntry(name).Open(); xml.Save(output); }
            var book = Read("xl/workbook.xml");
            var references = new XElement(ns + "externalReferences", new XElement(ns + "externalReference", new XAttribute(relationship + "id", "rIdExternal")));
            var calc = book.Root!.Element(ns + "calcPr"); if (calc != null) calc.AddBeforeSelf(references); else book.Root.Add(references);
            Write("xl/workbook.xml", book);
            var rels = Read("xl/_rels/workbook.xml.rels");
            rels.Root!.Add(new XElement(rels.Root.Name.Namespace + "Relationship", new XAttribute("Id", "rIdExternal"), new XAttribute("Type", relationship.NamespaceName + "/externalLink"), new XAttribute("Target", "externalLinks/externalLink1.xml")));
            Write("xl/_rels/workbook.xml.rels", rels);
            Write("xl/externalLinks/externalLink1.xml", new XDocument(new XElement(ns + "externalLink", new XElement(ns + "externalBook", new XAttribute(relationship + "id", "rId1"), new XElement(ns + "sheetNames", new XElement(ns + "sheetName", new XAttribute("val", "Rates")))))));
            XNamespace package = "http://schemas.openxmlformats.org/package/2006/relationships";
            Write("xl/externalLinks/_rels/externalLink1.xml.rels", new XDocument(new XElement(package + "Relationships", new XElement(package + "Relationship", new XAttribute("Id", "rId1"), new XAttribute("Type", relationship.NamespaceName + "/externalLinkPath"), new XAttribute("Target", "file:///Z:/missing-prices.xlsx"), new XAttribute("TargetMode", "External")))));
            var types = Read("[Content_Types].xml"); types.Root!.Add(new XElement(types.Root.Name.Namespace + "Override", new XAttribute("PartName", "/xl/externalLinks/externalLink1.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"))); Write("[Content_Types].xml", types);
            var lookup = Read("xl/worksheets/sheet2.xml"); var rows = lookup.Descendants(ns + "row").ToList();
            rows[0].Add(new XElement(ns + "c", new XAttribute("r", "B1"), new XAttribute("t", "inlineStr"), new XElement(ns + "is", new XElement(ns + "t", "external_price"))));
            rows[1].Add(new XElement(ns + "c", new XAttribute("r", "B2"), new XElement(ns + "f", "'[1]Rates'!A1"), new XElement(ns + "v", 99))); Write("xl/worksheets/sheet2.xml", lookup);
        }
        return path;
    }
    private static Dictionary<string, byte[]> Entries(string path)
    {
        using var zip = ZipFile.OpenRead(path);
        return zip.Entries.ToDictionary(e => e.FullName, e => { using var stream = e.Open(); using var buffer = new MemoryStream(); stream.CopyTo(buffer); return buffer.ToArray(); });
    }
    private static void CsvRoundTrip()
    {
        string path = Csv("csv"); var doc = DocumentStore.Open(path); var sheet = doc.Sheets[0];
        Assert(sheet.Cell(1, 1).Text == "  木剑  ", "Whitespace was trimmed.");
        doc.SetCells(sheet, new[] { (1, 1, "  名称,\"多行\"\n内容  "), (2, 1, "=SUM(A1:A2)") });
        string backup = DocumentStore.Save(doc); var reopened = DocumentStore.Open(path);
        Assert(reopened.Sheets[0].Cell(1, 1).Text == "  名称,\"多行\"\n内容  ", "CSV quoting changed data.");
        Assert(reopened.Sheets[0].Cell(2, 1).Text == "=SUM(A1:A2)" && !reopened.Sheets[0].Cell(2, 1).IsFormula, "CSV interpreted a formula.");
        Assert(reopened.Sheets[0].Cell(1, 0).Text == "001", "Leading zero lost.");
        Assert(File.ReadAllBytes(path).Take(3).SequenceEqual(new byte[] { 239, 187, 191 }), "BOM lost.");
        Assert(Directory.GetFiles(backup).Length > 0 && !doc.IsDirty, "Backup or saved revision missing.");
        string semi = Path.Combine(Output, "semicolon.csv"); File.WriteAllText(semi, "id;name\n1;  name  \n", new UTF8Encoding(false));
        var delimiter = DocumentStore.Open(semi); delimiter.SetCells(delimiter.Sheets[0], new[] { (1, 1, "edited") }); DocumentStore.Save(delimiter);
        Assert(File.ReadAllText(semi).StartsWith("id;name\n"), "CSV delimiter or newline changed.");
    }
    private static void ColumnHistory()
    {
        string path = Csv("columns");
        File.WriteAllText(path + ".schema.json", "{\"table_name\":\"test\",\"custom_root\":7,\"columns\":[{\"col_name\":\"price\",\"value_type\":\"int\",\"custom_column\":8,\"col_properties\":[{\"property_name\":\"cs_type\",\"property_value\":\"C\"}]}]}");
        var doc = DocumentStore.Open(path); var sheet = doc.Sheets[0]; var column = sheet.Columns[2];
        doc.RenameColumn(sheet, 2, "cost"); Assert(ReferenceEquals(sheet.Columns[2], column) && column.ColProperties![0].PropertyValue == "C", "Rename lost metadata.");
        doc.AddColumn(sheet, Column("enabled")); doc.SetCells(sheet, new[] { (1, 3, "true") });
        doc.DeleteColumn(sheet, 1); doc.Undo(); Assert(sheet.Cell(1, 3).Text == "true", "Column undo lost values.");
        doc.Undo(); doc.Undo(); doc.Undo(); Assert(!doc.IsDirty && column.ColName == "price", "Undo did not reach original state.");
        doc.Redo(); DocumentStore.Save(doc); doc.Undo(); Assert(doc.IsDirty && column.ColName == "price", "Undo after save should be dirty."); doc.Redo();
        using var json = JsonDocument.Parse(File.ReadAllText(path + ".schema.json"));
        Assert(json.RootElement.GetProperty("custom_root").GetInt32() == 7, "Unknown root metadata lost.");
        Assert(json.RootElement.GetProperty("columns").EnumerateArray().First(c => c.GetProperty("col_name").GetString() == "cost").GetProperty("custom_column").GetInt32() == 8, "Unknown column metadata lost.");
    }
    private static void XlsxReadOnlyAndSchema()
    {
        string path = Xlsx("xlsx"); byte[] original = File.ReadAllBytes(path); var doc = DocumentStore.Open(path); var sheet = doc.Sheets[0];
        Assert(doc.HasExternalLinks && sheet.Cell(1, 4).IsFormula && doc.ReadOnlyReason != null, "Formula, external link, or read-only detection missing.");
        Assert(sheet.Cell(1, 4).Text == "=C2*D2" && sheet.Cell(1, 4).Display == "240", "Formula expression or cached value missing.");
        Assert(doc.Sheets[1].Cell(1, 0).Text == "=Items!C2", "Cross-sheet formula missing.");
        Throws(() => doc.SetCells(sheet, new[] { (1, 2, "777") }), "XLSX cell edit should be rejected.");
        Throws(() => doc.AddColumn(sheet, Column("new_field")), "XLSX add column should be rejected.");
        Throws(() => doc.AddRow(sheet), "XLSX add row should be rejected.");
        Throws(() => doc.RenameColumn(sheet, 1, "title"), "XLSX rename column should be rejected.");
        doc.Execute("修改列类型", () => sheet.Columns[2].ValueType = "double", () => sheet.Columns[2].ValueType = "int");
        DocumentStore.Save(doc);
        Assert(File.ReadAllBytes(path).SequenceEqual(original), "XLSX source workbook was modified.");
        var reopened = DocumentStore.Open(path);
        Assert(reopened.Sheets[0].Columns[2].ValueType == "double", "XLSX schema update was not saved.");
    }
    private static void FormulaProtection()
    {
        string path = Xlsx("protected"); var doc = DocumentStore.Open(path); var sheet = doc.Sheets[0];
        Throws(() => doc.SetCells(sheet, new[] { (1, 1, "changed"), (1, 4, "0") }), "Protected paste was not rejected.");
        Assert(sheet.Cell(1, 1).Text != "changed" && !doc.IsDirty, "Rejected paste partially applied.");
        Throws(() => doc.SetCells(sheet, new[] { (1, 4, "=SUM(C2:D2)") }), "Formula expressions were editable.");
        Assert(sheet.ProtectionAt(1, 4) != null, "Formula protection message missing.");
        string csvPath = Csv("csv-literal"); var csvDoc = DocumentStore.Open(csvPath);
        csvDoc.SetCells(csvDoc.Sheets[0], new[] { (1, 1, "=text") });
        Assert(!csvDoc.Sheets[0].Cell(1, 1).IsFormula && csvDoc.Sheets[0].Cell(1, 1).Text == "=text", "CSV text entry created a formula.");
    }
    private static void ConflictProtection()
    {
        string path = Csv("conflict"); var doc = DocumentStore.Open(path); doc.SetCells(doc.Sheets[0], new[] { (1, 1, "mine") });
        File.AppendAllText(path, "003,external,1\r\n"); byte[] external = File.ReadAllBytes(path);
        Throws(() => DocumentStore.Save(doc), "Source conflict was overwritten."); Assert(File.ReadAllBytes(path).SequenceEqual(external), "External source changed.");
        doc = DocumentStore.Open(path); File.WriteAllText(path + ".schema.json", "{\"external\":true}");
        Throws(() => DocumentStore.Save(doc), "Schema conflict was overwritten.");
        Assert(File.ReadAllText(path + ".schema.json") == "{\"external\":true}", "External schema changed.");
    }
    private static void SchemaOnly()
    {
        string path = Xlsx("schema-only"); byte[] original = File.ReadAllBytes(path); var doc = DocumentStore.Open(path);
        var column = doc.Sheets[0].Columns[0]; string? before = column.ValueType;
        doc.Execute("type", () => column.ValueType = "int", () => column.ValueType = before);
        DocumentStore.Save(doc); Assert(File.ReadAllBytes(path).SequenceEqual(original), "Schema save rewrote workbook.");
    }
    private static void Rollback()
    {
        string path = Csv("rollback"); byte[] original = File.ReadAllBytes(path); var doc = DocumentStore.Open(path);
        doc.SetCells(doc.Sheets[0], new[] { (1, 1, "mine") });
        Directory.CreateDirectory(path + ".schema.json");
        Throws(() => DocumentStore.Save(doc), "Save should fail when schema target is a directory.");
        Assert(File.ReadAllBytes(path).SequenceEqual(original), "Source was not rolled back.");
        Assert(doc.IsDirty && doc.CanUndo, "Failed save discarded edit history.");
    }
    private static void Legacy()
    {
        string path = Path.Combine(Output, "legacy.xls");
        using (var workbook = new HSSFWorkbook()) { var sheet = workbook.CreateSheet("Data"); sheet.CreateRow(0).CreateCell(0).SetCellValue("id"); sheet.CreateRow(1).CreateCell(0).SetCellValue(1); using var stream = File.Create(path); workbook.Write(stream); }
        var doc = DocumentStore.Open(path); byte[] source = File.ReadAllBytes(path);
        Throws(() => doc.SetCells(doc.Sheets[0], new[] { (1, 0, "2") }), "Legacy XLS was editable.");
        doc.Sheets[0].Columns[0].ValueType = "int"; DocumentStore.Save(doc);
        Assert(File.ReadAllBytes(path).SequenceEqual(source), "XLS metadata save modified data.");
    }
    private static void Validation()
    {
        var doc = DocumentStore.Open(Csv("validation")); var sheet = doc.Sheets[0];
        sheet.Columns[0].PkOrder = 0; sheet.Columns[2].ValueType = "int";
        doc.SetCells(sheet, new[] { (2, 0, "001"), (1, 2, "bad") });
        var errors = DocumentValidation.Validate(sheet);
        Assert(errors.Any(e => e.Row == 2 && e.Column == "price") && errors.Any(e => e.Message.Contains("主键")), "Validation ignored unsaved changes.");
    }
    private static void Ui()
    {
        using var watchdog = new System.Threading.Timer(_ => { Console.Error.WriteLine("UI check timed out; a modal dialog or layout event did not finish."); Environment.Exit(2); }, null, TimeSpan.FromSeconds(35), System.Threading.Timeout.InfiniteTimeSpan);
        ApplicationConfiguration.Initialize();
        string directory = Path.Combine(Output, "preview"); Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "Items.xlsx"); File.Copy(Xlsx("preview"), path, true);
        string examples = Path.GetFullPath(Path.Combine(Output, "../../Examples"));
        File.Copy(Path.Combine(examples, "schema.registry.json"), Path.Combine(directory, "schema.registry.json"), true);
        File.Copy(Path.Combine(examples, "Items.csv"), Path.Combine(directory, "Items.csv"), true);
        using var form = new MainForm(path) { ShowInTaskbar = false, StartPosition = FormStartPosition.Manual, Location = new Point(-20000, -20000) };
        form.Show(); Application.DoEvents();
        Console.WriteLine("UI: window opened");
        var view = Descendants(form).OfType<TableEditorControl>().First();
        Assert(view.DataGrid.ColumnCount == 5 && view.DataGrid.RowCount == 8, "Data grid dimensions are wrong.");
        view.DataGrid.CurrentCell = view.DataGrid[1, 0]; Assert(!view.DataGrid.BeginEdit(false), "Excel data should be read-only.");
        view.DataGrid.CurrentCell = view.DataGrid[4, 0]; Assert(!view.DataGrid.BeginEdit(false), "UI formula lock failed.");
        Assert(view.ProtectionMessage?.Contains("公式") == true && Descendants(view).Any(c => c.Name == "protected_help" && c.Visible), "Protected cell lacks a lock explanation button.");
        Assert(view.DataGrid[4, 0].Value?.ToString() == "240", "Grid displayed formula text instead of the saved calculation result.");
        Console.WriteLine("UI: excel read-only and formula cached result verified");
        string csvPath = Path.Combine(directory, "Items.csv");
        form.OpenFile(csvPath); Application.DoEvents();
        var csvView = Descendants(form).OfType<TableEditorControl>().First(v => v.Document.IsCsv);
        csvView.DataGrid.CurrentCell = csvView.DataGrid[1, 0]; csvView.DataGrid.BeginEdit(false);
        ((TextBox)csvView.DataGrid.EditingControl!).Text = "UI edit"; Assert(csvView.CommitEdit(), "CSV grid commit failed."); Application.DoEvents();
        Assert(csvView.Sheet.Cell(1, 1).Text == "UI edit", "CSV virtual grid lost edits.");
        csvView.Document.JumpTo(0); Application.DoEvents();
        Assert(!csvView.Document.IsDirty && csvView.Sheet.Cell(1, 1).Text != "UI edit", "History jump did not restore CSV.");
        Console.WriteLine("UI: csv edit and history jump verified");
        form.OpenFile(path, "Lookup"); Application.DoEvents();
        Assert(Descendants(form).OfType<TableEditorControl>().Where(v => !v.Document.IsCsv).All(v => ReferenceEquals(v.Document, view.Document)), "Sheets do not share the same workbook.");
        form.OpenFile(path, "Items"); Application.DoEvents();
        view.DataGrid.CurrentCell = view.DataGrid[2, 0]; Application.DoEvents();
        form.Size = new Size(1060, 700); Application.DoEvents();
        Assert(Descendants(form).OfType<SplitContainer>().All(s => s.Panel1.Width > 0 && s.Panel2.Width > 0), "Minimum window size hides panels.");
        form.ClientSize = new Size(1440, 860); Application.DoEvents();
        view.DataGrid.Focus(); Application.DoEvents();
        view.DataGrid.CurrentCell = view.DataGrid[4, 0]; Application.DoEvents();
        Assert(view.SelectedColumn == 4 && view.ProtectionMessage != null, "Protected state did not follow cell selection after resizing.");
        using var bitmap = new Bitmap(form.Width, form.Height); form.DrawToBitmap(bitmap, new Rectangle(Point.Empty, bitmap.Size));
        bitmap.Save(Path.Combine(Output, "studio-preview.png"));
        Console.WriteLine("UI: screenshot saved");
        form.Close();
    }
    private static void AutoInferredTypes()
    {
        string path = Path.Combine(Output, "auto-types-" + Guid.NewGuid().ToString("N")[..6] + ".csv");
        File.WriteAllText(path, "id,code,name,price,rate,enabled,tags\r\n1,001,木剑,100,1.5,true,\"[1,2,3]\"\r\n2,002,盾牌,200,2.25,false,\"[4,5]\"\r\n", new UTF8Encoding(false));
        var registry = new PropertyRegistry { Properties = new() { new() { Key = "cs_type", Title = "终端类型", Scope = "column", Kind = "enum", Default = "ALL", EnumValues = new() { "ALL", "C", "S" } } } };
        var doc = DocumentStore.Open(path); var sheet = doc.Sheets[0];
        Assert(sheet.Columns[0].ValueType == "int" && sheet.Columns[0].IsConfigured, "int column was not auto-inferred.");
        Assert(sheet.Columns[1].ValueType == "string" && sheet.Columns[1].IsConfigured, "Leading-zero code column should stay string.");
        Assert(sheet.Columns[2].ValueType == "string" && sheet.Columns[2].IsConfigured, "string column was not auto-inferred.");
        Assert(sheet.Columns[3].ValueType == "int" && sheet.Columns[3].IsConfigured, "price column was not auto-inferred as int.");
        Assert(sheet.Columns[4].ValueType == "double" && sheet.Columns[4].IsConfigured, "decimal column was not auto-inferred as double.");
        Assert(sheet.Columns[5].ValueType == "bool" && sheet.Columns[5].IsConfigured, "bool column was not auto-inferred.");
        Assert(sheet.Columns[6].ValueType == "int[]" && sheet.Columns[6].IsConfigured, "int[] column was not auto-inferred.");
        Assert(sheet.Columns.All(c => ColumnConfiguration.Missing(c, registry) == null && (c.ColProperties == null || c.ColProperties.Count == 0)), "Auto-inferred columns should be ready without forcing optional custom properties.");
        DocumentStore.Save(doc);
        var reopened = DocumentStore.Open(path);
        Assert(reopened.Sheets[0].Columns.Count == 7 && reopened.Sheets[0].Columns.All(c => c.HasStoredDefinition && c.IsConfigured), "Saving did not persist all auto-inferred column types.");
        doc.AddColumn(sheet, Column("optional")); DocumentStore.Save(doc);
        var afterAdd = DocumentStore.Open(path);
        Assert(afterAdd.Sheets[0].Columns[7].IsConfigured && afterAdd.Sheets[0].Columns[7].ColProperties?.Count == 0, "A typed column without attributes could not be saved.");
        Throws(() => doc.AddColumn(sheet, new ColumnSchema { ColName = "unconfigured" }), "New column accepted without type confirmation.");
    }
    private static void WorkspaceSchemaSync()
    {
        string dir = Path.Combine(Output, "ws-sync-" + Guid.NewGuid().ToString("N")[..6]);
        Directory.CreateDirectory(dir);
        string csvPath = Path.Combine(dir, "Heroes.csv");
        File.WriteAllText(csvPath, "id,name,hp,is_boss\r\n101,亚瑟,1500,false\r\n102,魔王,9999,true\r\n", new UTF8Encoding(false));
        byte[] csvBytes = File.ReadAllBytes(csvPath);
        string partialPath = Path.Combine(dir, "Partial.csv");
        File.WriteAllText(partialPath, "id,title,score\r\n1,一,9.5\r\n2,二,8.2\r\n", new UTF8Encoding(false));
        File.WriteAllText(partialPath + ".schema.json", "{\"table_name\":\"Partial.csv\",\"columns\":[{\"col_name\":\"id\",\"value_type\":\"long\",\"col_properties\":[{\"property_name\":\"cs_type\",\"property_value\":\"S\"}]}]}");
        string badCsv = Path.Combine(dir, "Broken.csv");
        File.WriteAllText(badCsv, ",dup,dup\r\n1,2,3\r\n", new UTF8Encoding(false));

        var first = DocumentStore.EnsureDirectorySchemas(dir);
        Assert(first.Created == 1 && first.Updated == 1 && first.Errors.Count == 0, $"Unexpected first sync counts: created={first.Created}, updated={first.Updated}, errors={first.Errors.Count}");
        Assert(File.ReadAllBytes(csvPath).SequenceEqual(csvBytes), "Workspace schema sync modified source CSV.");
        Assert(File.Exists(csvPath + ".schema.json") && !File.Exists(badCsv + ".schema.json"), "Schema file creation mismatch.");
        var heroes = DocumentStore.Open(csvPath).Sheets[0];
        Assert(heroes.Columns[0].ValueType == "int" && heroes.Columns[1].ValueType == "string" && heroes.Columns[2].ValueType == "int" && heroes.Columns[3].ValueType == "bool", "Heroes auto-inferred types mismatch.");
        var partial = DocumentStore.Open(partialPath).Sheets[0];
        Assert(partial.Columns[0].ValueType == "long" && partial.Columns[0].ColProperties![0].PropertyValue == "S", "Existing column definition was overwritten during sync.");
        Assert(partial.Columns[1].ValueType == "string" && partial.Columns[2].ValueType == "double", "Newly added columns were not supplemented.");
        var second = DocumentStore.EnsureDirectorySchemas(dir);
        Assert(second.Created == 0 && second.Updated == 0, "Workspace schema sync was not idempotent.");
    }
    private static void HistoryJumpAndLimit()
    {
        var doc = DocumentStore.Open(Csv("history"));
        var sheet = doc.Sheets[0];
        string original = sheet.Cell(1, 1).Text;
        for (int i = 1; i <= 5; i++) doc.SetCells(sheet, new[] { (1, 1, $"v{i}") });
        Assert(doc.HistoryCount == 5 && doc.HistoryCursor == 5 && doc.GetHistory().Count == 6, "History count mismatch.");
        doc.JumpTo(2);
        Assert(sheet.Cell(1, 1).Text == "v2" && doc.HistoryCursor == 2 && doc.CanUndo && doc.CanRedo, "Jump backward failed.");
        doc.JumpTo(4);
        Assert(sheet.Cell(1, 1).Text == "v4" && doc.HistoryCursor == 4, "Jump forward failed.");
        doc.JumpTo(0);
        Assert(sheet.Cell(1, 1).Text == original && !doc.IsDirty && !doc.CanUndo && doc.CanRedo, "Jump to initial state failed.");
        doc.JumpTo(2);
        doc.SetCells(sheet, new[] { (1, 1, "branch") });
        Assert(doc.HistoryCount == 3 && doc.HistoryCursor == 3 && !doc.CanRedo && sheet.Cell(1, 1).Text == "branch", "Branching after jump did not trim redo tail.");
        for (int i = 1; i <= 120; i++) doc.SetCells(sheet, new[] { (1, 1, $"step-{i}") });
        Assert(doc.HistoryCount == EditorDocument.MaxHistory && doc.HistoryCursor == EditorDocument.MaxHistory && doc.HistoryTrimmed, "History did not cap at 100 entries.");
        doc.JumpTo(0);
        Assert(sheet.Cell(1, 1).Text == "step-20" && doc.HistoryCursor == 0, "Jump to earliest retained state after 100-entry trim failed.");
    }
    private static void SpecialRanges()
    {
        string path = Xlsx("ranges");
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            var entry = zip.GetEntry("xl/worksheets/sheet1.xml")!;
            XDocument xml; using (var stream = entry.Open()) xml = XDocument.Load(stream);
            XElement Cell(string address) => xml.Descendants(ns + "c").First(c => (string?)c.Attribute("r") == address);
            var array = Cell("E2").Element(ns + "f")!; array.SetAttributeValue("t", "array"); array.SetAttributeValue("ref", "E2:E3");
            var shared = Cell("E4").Element(ns + "f")!; shared.SetAttributeValue("t", "shared"); shared.SetAttributeValue("si", "0"); shared.SetAttributeValue("ref", "E4:E5");
            var follower = Cell("E5").Element(ns + "f")!; follower.Value = ""; follower.SetAttributeValue("t", "shared"); follower.SetAttributeValue("si", "0");
            Cell("E6").Element(ns + "f")!.Value = "'[1]Lookup'!A2";
            xml.Root!.Add(new XElement(ns + "mergeCells", new XElement(ns + "mergeCell", new XAttribute("ref", "A7:B7"))));
            entry.Delete(); using var output = zip.CreateEntry("xl/worksheets/sheet1.xml").Open(); xml.Save(output);
        }
        byte[] original = File.ReadAllBytes(path);
        var doc = DocumentStore.Open(path); var sheet = doc.Sheets[0];
        foreach (var p in new[] { (1, 4), (2, 4), (3, 4), (4, 4), (5, 4), (6, 1) })
        {
            Assert(sheet.ProtectionAt(p.Item1, p.Item2) != null, "Special formula or merged cell lost protection metadata.");
            Throws(() => doc.SetCells(sheet, new[] { (p.Item1, p.Item2, "0") }), "Special formula or merged cell was editable.");
        }
        DocumentStore.Save(doc);
        Assert(File.ReadAllBytes(path).SequenceEqual(original) && DocumentStore.Open(path).Sheets[0].Cell(1, 4).Text == "=C2*D2", "Array formula workbook was altered.");
    }
    private static void MixedSheets()
    {
        string path = Path.Combine(Output, "mixed-" + Guid.NewGuid().ToString("N")[..6] + ".xlsx");
        using (var workbook = new XSSFWorkbook())
        {
            var good = workbook.CreateSheet("Config"); good.CreateRow(0).CreateCell(0).SetCellValue("id"); good.CreateRow(1).CreateCell(0).SetCellValue(1);
            var auxiliary = workbook.CreateSheet("Auxiliary"); auxiliary.CreateRow(1).CreateCell(0).SetCellValue("not a config header");
            workbook.CreateSheet("Empty"); using var output = File.Create(path); workbook.Write(output);
        }
        byte[] original = File.ReadAllBytes(path);
        var doc = DocumentStore.Open(path);
        Assert(doc.Sheets.Count == 3 && doc.Sheets[1].SchemaIssue != null, "Auxiliary sheet blocked the workbook or was treated as a valid config.");
        DocumentStore.Save(doc);
        Assert(File.ReadAllBytes(path).SequenceEqual(original), "Mixed workbook was altered during schema save.");
        using var json = JsonDocument.Parse(File.ReadAllText(path + ".schema.json"));
        Assert(json.RootElement.GetProperty("sheets").GetArrayLength() == 1, "Auxiliary or empty sheet emitted invalid schema.");
    }
    private static IEnumerable<Control> Descendants(Control root)
    {
        foreach (Control child in root.Controls) { yield return child; foreach (Control nested in Descendants(child)) yield return nested; }
    }
}
