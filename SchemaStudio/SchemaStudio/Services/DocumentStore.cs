using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using CsvHelper;
using CsvHelper.Configuration;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using SchemaStudio.Models;

namespace SchemaStudio.Services;

public static class DocumentStore
{
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace Relations = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public static EditorDocument Open(string path)
    {
        path = Path.GetFullPath(path);
        var doc = new EditorDocument { Path = path, SourceBytes = File.ReadAllBytes(path) };
        string schemaPath = DataIO.GetSchemaJsonPath(path);
        doc.SchemaBytes = File.Exists(schemaPath) ? File.ReadAllBytes(schemaPath) : null;
        if (doc.IsCsv) ReadCsv(doc);
        else if (doc.IsLegacy) ReadLegacy(doc);
        else if (Path.GetExtension(path).Equals(".xlsx", StringComparison.OrdinalIgnoreCase)) ReadXlsx(doc);
        else throw new InvalidOperationException("支持 CSV、XLSX 和 XLS 文件。");
        ApplySchema(doc);
        foreach (var sheet in doc.Sheets)
        {
            sheet.OriginalCells = new(sheet.Cells, StringComparer.OrdinalIgnoreCase);
            sheet.OriginalRowCount = sheet.RowCount;
            sheet.OriginalColumnCount = sheet.Columns.Count;
        }
        return doc;
    }

    private static void ReadCsv(EditorDocument doc)
    {
        doc.Encoding = DataIO.DetectEncoding(doc.Path);
        using var stream = new MemoryStream(doc.SourceBytes);
        using var reader = new StreamReader(stream, doc.Encoding, true);
        string content = reader.ReadToEnd();
        doc.NewLine = content.Contains("\r\n") ? "\r\n" : content.Contains('\n') ? "\n" : content.Contains('\r') ? "\r" : Environment.NewLine;
        var config = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            HasHeaderRecord = false, DetectDelimiter = true, IgnoreBlankLines = false,
            DetectDelimiterValues = new[] { ",", ";", "\t", "|" }
        };
        using var csv = new CsvReader(new StringReader(content), config);
        var sheet = new EditorSheet { Name = Path.GetFileNameWithoutExtension(doc.Path) };
        int row = 0, width = 0;
        while (csv.Read())
        {
            width = Math.Max(width, csv.Parser.Count);
            for (int col = 0; col < csv.Parser.Count; col++) sheet.Cells[EditorSheet.Address(row, col)] = new(csv.GetField(col) ?? "");
            row++;
        }
        doc.Delimiter = csv.Parser.Delimiter;
        sheet.RowCount = Math.Max(0, row - 1);
        BuildColumns(sheet, width);
        doc.Sheets.Add(sheet);
    }

    private static void ReadXlsx(EditorDocument doc)
    {
        using var stream = new MemoryStream(doc.SourceBytes);
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read);
        var workbook = ReadXml(zip, "xl/workbook.xml");
        if (workbook.Root?.Name.Namespace != Main) throw new InvalidOperationException("暂不支持此 Excel XML 格式。请用 Excel 另存为标准 XLSX。");
        var rels = ReadXml(zip, "xl/_rels/workbook.xml.rels").Root!.Elements()
            .Where(x => (string?)x.Attribute("TargetMode") != "External")
            .ToDictionary(x => (string)x.Attribute("Id")!, x => ResolvePart("xl/workbook.xml", (string)x.Attribute("Target")!));
        var strings = zip.GetEntry("xl/sharedStrings.xml") == null ? new List<string>() :
            ReadXml(zip, "xl/sharedStrings.xml").Root!.Elements(Main + "si").Select(ReadText).ToList();
        doc.HasExternalLinks = zip.Entries.Any(e => e.FullName.StartsWith("xl/externalLinks/", StringComparison.OrdinalIgnoreCase)) ||
            workbook.Descendants(Main + "externalReference").Any();
        if (doc.HasExternalLinks) doc.Notices.Add("包含外部工作簿链接：保留链接和缓存，不联网更新；依赖结果需在 Excel 中重算。");
        doc.ReadOnlyReason = zip.Entries.Any(e => e.FullName.StartsWith("_xmlsignatures/", StringComparison.OrdinalIgnoreCase))
            ? "此工作簿有数字签名且 Excel 数据为只读；仅可查看数据和配置 .schema.json 属性。"
            : "Excel 数据为只读（避免破坏公式、样式和链接）；可在此配置字段类型、主键与属性，需改数据时请点击「外部打开」在 Excel 中编辑后重新加载。";
        foreach (var element in workbook.Descendants(Main + "sheet"))
        {
            string name = (string)element.Attribute("name")!;
            string part = rels[(string)element.Attribute(Relations + "id")!];
            if (!part.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase)) continue;
            var xml = ReadXml(zip, part);
            var sheet = new EditorSheet { Name = name, PartPath = part,
                HasTable = xml.Descendants(Main + "tablePart").Any(), IsProtected = xml.Root!.Element(Main + "sheetProtection") != null };
            int width = 0;
            foreach (var cell in xml.Descendants(Main + "sheetData").Elements(Main + "row").Elements(Main + "c"))
            {
                string? address = (string?)cell.Attribute("r");
                if (address == null) throw new InvalidOperationException("工作簿存在没有地址的单元格，请在 Excel 中重新保存。");
                var formula = cell.Element(Main + "f");
                string type = (string?)cell.Attribute("t") ?? "n";
                string raw = cell.Element(Main + "v")?.Value ?? "";
                string value = type switch
                {
                    "s" => int.TryParse(raw, out int i) && i >= 0 && i < strings.Count ? strings[i] : throw new InvalidOperationException("Excel 共享字符串索引无效。"),
                    "inlineStr" => ReadText(cell.Element(Main + "is")),
                    "b" => raw == "1" ? "true" : "false",
                    _ => raw
                };
                if (formula == null && value.Length == 0) continue;
                string? protection = null;
                if (formula != null)
                {
                    string? formulaType = (string?)formula.Attribute("t");
                    if (formulaType is "shared" or "array" or "dataTable") protection = "共享、数组及数据表公式由 Excel 维护。";
                    if (formula.Value.Contains('[') || formula.Value.Contains(']')) protection = "外部链接或结构化引用公式由 Excel 维护。";
                    if ((string?)formula.Attribute("ref") is { } range) sheet.ProtectedRanges.Add(range);
                }
                if (cell.Attribute("cm") != null || cell.Attribute("vm") != null) protection = "此单元格包含动态数组或扩展元数据，请在 Excel 中编辑。";
                var position = EditorSheet.Position(address);
                sheet.RowCount = Math.Max(sheet.RowCount, position.Row);
                width = Math.Max(width, position.Column + 1);
                sheet.Cells[address] = formula != null ? new("=" + formula.Value, value, "formula", protection, cell.Element(Main + "v") != null && (raw.Length > 0 || type == "str")) :
                    new(value, "", type == "n" ? "number" : type == "b" ? "boolean" : "string", protection);
            }
            foreach (var merge in xml.Descendants(Main + "mergeCell"))
                if ((string?)merge.Attribute("ref") is { } range) sheet.ProtectedRanges.Add(range);
            if (width > 16_384 || sheet.RowCount >= 1_048_576) throw new InvalidOperationException("工作表超出支持的行列范围。");
            BuildColumns(sheet, width);
            doc.Sheets.Add(sheet);
        }
        if (doc.Sheets.Count == 0) throw new InvalidOperationException("没有可以编辑的工作表。");
        doc.Notices.Add("公式显示文件中的上次保存结果，不在本工具内重算。请在 Excel / WPS 中编辑、重算并保存后重新加载。");
    }

    private static void ReadLegacy(EditorDocument doc)
    {
        doc.ReadOnlyReason = "旧版 XLS 仅支持查看和属性配置；请在 Excel 中另存为 XLSX 后编辑数据，以保留公式和链接。";
        using var stream = new MemoryStream(doc.SourceBytes);
        using var workbook = new HSSFWorkbook(stream);
        var format = new DataFormatter(CultureInfo.InvariantCulture);
        for (int i = 0; i < workbook.NumberOfSheets; i++)
        {
            var source = workbook.GetSheetAt(i);
            var sheet = new EditorSheet { Name = source.SheetName, RowCount = source.LastRowNum };
            int width = 0;
            foreach (IRow row in source)
                foreach (var cell in row.Cells)
                {
                    width = Math.Max(width, cell.ColumnIndex + 1);
                    sheet.Cells[EditorSheet.Address(row.RowNum, cell.ColumnIndex)] = cell.CellType == CellType.Formula
                        ? new("=" + cell.CellFormula, cell.CachedFormulaResultType switch
                        {
                            CellType.Numeric => cell.NumericCellValue.ToString("G17", CultureInfo.InvariantCulture),
                            CellType.String => cell.StringCellValue,
                            CellType.Boolean => cell.BooleanCellValue ? "true" : "false",
                            CellType.Error => "#ERROR",
                            _ => ""
                        }, "formula") : new(format.FormatCellValue(cell));
                }
            BuildColumns(sheet, width); doc.Sheets.Add(sheet);
        }
        doc.Notices.Add(doc.ReadOnlyReason);
    }

    private static void BuildColumns(EditorSheet sheet, int width)
    {
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (int col = 0; col < width; col++)
        {
            string name = sheet.Cell(0, col).Display.Trim();
            if (string.IsNullOrWhiteSpace(name))
            {
                sheet.SchemaIssue = "第 1 行存在空字段名，此 Sheet 仅供查看。请在 Excel 中整理表头后重新打开。";
                name = $"column_{col + 1}";
                while (names.Contains(name)) name += "_";
            }
            if (!names.Add(name))
            {
                sheet.SchemaIssue = "第 1 行存在重复字段名，此 Sheet 仅供查看。请在源文件中整理表头后重新打开。";
                do { name += "_"; } while (!names.Add(name));
            }
            sheet.Columns.Add(new ColumnSchema { ColName = name, ValueType = "string", ColSort = col, ColProperties = new() });
        }
        if (sheet.SchemaIssue != null) sheet.IsProtected = true;
    }

    private static void ApplySchema(EditorDocument doc)
    {
        TableSchema? table = null;
        if (doc.SchemaBytes != null)
        {
            using var reader = new StreamReader(new MemoryStream(doc.SchemaBytes), Encoding.UTF8, true);
            string schemaText = reader.ReadToEnd();
            using var json = JsonDocument.Parse(schemaText);
            if (json.RootElement.EnumerateObject().Any(p => p.Name.Equals("sheets", StringComparison.OrdinalIgnoreCase)))
                doc.SchemaEnvelope = JsonSerializer.Deserialize<WorkbookSchema>(schemaText, DataIO.JsonOptions);
            else { table = JsonSerializer.Deserialize<TableSchema>(schemaText, DataIO.JsonOptions); doc.TableEnvelope = table; }
        }
        foreach (var sheet in doc.Sheets)
        {
            if (sheet.SchemaIssue != null) continue;
            var entry = doc.SchemaEnvelope?.Sheets.FirstOrDefault(s => s.SheetName == sheet.Name);
            var old = entry?.Columns ?? (doc.IsCsv || sheet == doc.Sheets[0] ? table?.Columns : null);
            sheet.Properties = entry?.SheetProperties ?? (sheet == doc.Sheets[0] ? table?.TableProperties : null) ?? new();
            for (int i = 0; i < sheet.Columns.Count; i++)
            {
                var column = old?.FirstOrDefault(c => string.Equals(c.ColName, sheet.Columns[i].ColName, StringComparison.OrdinalIgnoreCase));
                if (column == null)
                {
                    sheet.Columns[i].ValueType = TypeInference.InferColumnType(sheet, i, doc.IsCsv);
                    sheet.Columns[i].IsConfigured = true;
                    continue;
                }
                column.ColSort = i;
                if (string.IsNullOrWhiteSpace(column.ValueType))
                    column.ValueType = TypeInference.InferColumnType(sheet, i, doc.IsCsv);
                column.IsConfigured = true;
                column.HasStoredDefinition = true;
                if (column.PrimaryKey && column.PkOrder < 0) column.PkOrder = i;
                sheet.Columns[i] = column;
            }
        }
    }

    public static string SchemaJson(EditorDocument doc)
    {
        foreach (var sheet in doc.Sheets)
            for (int i = 0; i < sheet.Columns.Count; i++) { sheet.Columns[i].ColSort = i; sheet.Columns[i].PrimaryKey = sheet.Columns[i].PkOrder >= 0; }
        if (doc.IsCsv)
        {
            if (doc.Sheets[0].SchemaIssue != null) return doc.SchemaBytes == null ? "{}" : Encoding.UTF8.GetString(doc.SchemaBytes);
            var table = doc.TableEnvelope ?? new TableSchema();
            table.TableName = Path.GetFileName(doc.Path); table.Columns = doc.Sheets[0].Columns.Where(c => c.IsConfigured || c.HasStoredDefinition).ToList(); table.TableProperties = doc.Sheets[0].Properties;
            return JsonSerializer.Serialize(table, DataIO.JsonOptions);
        }
        var envelope = doc.SchemaEnvelope ?? new WorkbookSchema { WorkbookName = Path.GetFileName(doc.Path) };
        foreach (var sheet in doc.Sheets)
        {
            if (sheet.SchemaIssue != null || sheet.Columns.Count == 0) continue;
            var entry = envelope.Sheets.FirstOrDefault(s => s.SheetName == sheet.Name);
            if (entry == null) { entry = new SheetSchema { SheetName = sheet.Name }; envelope.Sheets.Add(entry); }
            entry.Columns = sheet.Columns.Where(c => c.IsConfigured || c.HasStoredDefinition).ToList(); entry.SheetProperties = sheet.Properties;
        }
        return JsonSerializer.Serialize(envelope, DataIO.JsonOptions);
    }

    public static bool NeedsSchemaSync(EditorDocument doc)
    {
        var validSheets = doc.Sheets.Where(s => s.SchemaIssue == null && s.Columns.Count > 0).ToList();
        if (validSheets.Count == 0) return false;
        if (doc.SchemaBytes == null) return true;
        return validSheets.Any(s =>
            s.Columns.Any(c => !c.HasStoredDefinition) ||
            (!doc.IsCsv && (doc.SchemaEnvelope == null || !doc.SchemaEnvelope.Sheets.Any(e => e.SheetName == s.Name))));
    }

    public static (int Created, int Updated, List<string> Errors) EnsureDirectorySchemas(string rootDirectory, IReadOnlySet<string>? skipPaths = null)
    {
        int created = 0, updated = 0;
        var errors = new List<string>();
        if (!Directory.Exists(rootDirectory)) return (0, 0, errors);

        var stack = new Stack<string>();
        stack.Push(rootDirectory);
        while (stack.Count > 0)
        {
            string dir = stack.Pop();
            IEnumerable<string> subDirs;
            try { subDirs = Directory.GetDirectories(dir); } catch { continue; }
            foreach (string sub in subDirs)
            {
                string name = Path.GetFileName(sub);
                if (name.StartsWith('.') || name is "bin" or "obj") continue;
                try { if ((File.GetAttributes(sub) & FileAttributes.ReparsePoint) != 0) continue; } catch { continue; }
                stack.Push(sub);
            }

            IEnumerable<string> files;
            try { files = Directory.GetFiles(dir); } catch { continue; }
            foreach (string file in files)
            {
                string fileName = Path.GetFileName(file);
                if (fileName.StartsWith("~$") || !UiDefaults.FileExts.Contains(Path.GetExtension(file).ToLowerInvariant())) continue;
                if (skipPaths != null && skipPaths.Contains(Path.GetFullPath(file))) continue;
                try
                {
                    var doc = Open(file);
                    if (!NeedsSchemaSync(doc)) continue;
                    bool isNew = doc.SchemaBytes == null;
                    Save(doc);
                    if (isNew) created++; else updated++;
                }
                catch (Exception ex)
                {
                    errors.Add($"{fileName}: {ex.Message}");
                }
            }
        }
        return (created, updated, errors);
    }

    public static bool HasDataChanges(EditorDocument doc) => doc.Sheets.Any(s => s.RowCount != s.OriginalRowCount || s.Columns.Count != s.OriginalColumnCount || s.Cells.Count != s.OriginalCells.Count ||
        s.Cells.Any(c => !s.OriginalCells.TryGetValue(c.Key, out var original) || c.Value != original));

    public static string Save(EditorDocument doc)
    {
        if (doc.IsCsv && doc.Sheets[0].SchemaIssue is { } issue) throw new InvalidOperationException(issue);
        string schemaPath = DataIO.GetSchemaJsonPath(doc.Path);
        CheckUnchanged(doc.Path, doc.SourceBytes); CheckUnchanged(schemaPath, doc.SchemaBytes);
        byte[] source = doc.SourceBytes;
        bool dataChanged = HasDataChanges(doc);
        if (dataChanged)
        {
            if (doc.ReadOnlyReason != null || !doc.IsCsv) throw new InvalidOperationException(doc.ReadOnlyReason ?? "Excel 数据为只读，仅保存 .schema.json 配置。");
            source = WriteCsv(doc);
        }
        byte[] schema = new UTF8Encoding(false).GetBytes(SchemaJson(doc));
        string backupDirectory = Path.Combine(Path.GetDirectoryName(doc.Path)!, ".schemastudio-backups");
        Directory.CreateDirectory(backupDirectory);
        string id = DateTime.Now.ToString("yyyyMMdd-HHmmss-fff") + "-" + Guid.NewGuid().ToString("N")[..6];
        var committed = new List<(string Path, byte[]? Original, byte[] Written)>();
        try
        {
            if (dataChanged)
            {
                Replace(doc.Path, source, doc.SourceBytes, Path.Combine(backupDirectory, Path.GetFileName(doc.Path) + "." + id + ".bak"));
                committed.Add((doc.Path, doc.SourceBytes, source));
            }
            Replace(schemaPath, schema, doc.SchemaBytes, Path.Combine(backupDirectory, Path.GetFileName(schemaPath) + "." + id + ".bak"));
            committed.Add((schemaPath, doc.SchemaBytes, schema));
        }
        catch (Exception error)
        {
            var rollbackErrors = new List<string>();
            foreach (var file in committed.AsEnumerable().Reverse())
                try
                {
                    CheckUnchanged(file.Path, file.Written);
                    if (file.Original == null) File.Delete(file.Path);
                    else Replace(file.Path, file.Original, file.Written, null);
                }
                catch (Exception rollback) { rollbackErrors.Add(rollback.Message); }
            throw new IOException(rollbackErrors.Count == 0 ? "保存失败，已恢复本次写入：" + error.Message :
                "保存失败，部分文件未能恢复。请从备份目录恢复：" + backupDirectory + "\n" + string.Join("\n", rollbackErrors), error);
        }
        doc.SourceBytes = source; doc.SchemaBytes = schema;
        foreach (var sheet in doc.Sheets)
        {
            sheet.OriginalCells = new(sheet.Cells, StringComparer.OrdinalIgnoreCase);
            sheet.OriginalRowCount = sheet.RowCount; sheet.OriginalColumnCount = sheet.Columns.Count;
        }
        doc.MarkSaved();
        return backupDirectory;
    }

    private static byte[] WriteCsv(EditorDocument doc)
    {
        using var stream = new MemoryStream();
        using (var writer = new StreamWriter(stream, doc.Encoding, leaveOpen: true))
        using (var csv = new CsvWriter(writer, new CsvConfiguration(CultureInfo.InvariantCulture) { Delimiter = doc.Delimiter, NewLine = doc.NewLine }))
        {
            var sheet = doc.Sheets[0];
            for (int row = 0; row <= sheet.RowCount; row++)
            {
                for (int col = 0; col < sheet.Columns.Count; col++) csv.WriteField(sheet.Cell(row, col).Text);
                csv.NextRecord();
            }
        }
        return stream.ToArray();
    }

    private static void CheckUnchanged(string path, byte[]? expected)
    {
        bool exists = File.Exists(path);
        if (exists != (expected != null) || exists && !File.ReadAllBytes(path).AsSpan().SequenceEqual(expected))
            throw new IOException($"文件已被外部修改或移动，未覆盖：{Path.GetFileName(path)}。请先备份当前修改，再重新打开文件。");
    }

    private static void Replace(string path, byte[] bytes, byte[]? expected, string? backup)
    {
        string temp = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            using (var output = new FileStream(temp, FileMode.CreateNew, FileAccess.Write)) { output.Write(bytes); output.Flush(true); }
            CheckUnchanged(path, expected);
            if (expected == null) File.Move(temp, path);
            else File.Replace(temp, path, backup);
        }
        finally { if (File.Exists(temp)) File.Delete(temp); }
    }

    private static string ResolvePart(string owner, string target) => new Uri(new Uri("https://package/" + owner), target).AbsolutePath.TrimStart('/');
    private static string ReadText(XElement? element) => element == null ? "" :
        System.Text.RegularExpressions.Regex.Replace(string.Concat(element.Descendants(Main + "t").Where(t => t.Parent?.Name != Main + "rPh").Select(t => t.Value)),
            "_x([0-9a-fA-F]{4})_", match => ((char)Convert.ToInt32(match.Groups[1].Value, 16)).ToString());
    private static XDocument ReadXml(ZipArchive zip, string path)
    {
        using var stream = (zip.GetEntry(path) ?? throw new InvalidOperationException("缺少工作簿部件：" + path)).Open();
        return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
    }
}
