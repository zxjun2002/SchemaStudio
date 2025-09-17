// Program.cs — SchemaStudio（Newtonsoft.Json + NPOI，多 Sheet，懒加载 Sheet 树）
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using System.Windows.Forms;
using CsvHelper;
using ExcelDataReader;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using NPOI.SS.UserModel;     // 接口类型
using NPOI.HSSF.UserModel;   // .xls
using NPOI.XSSF.UserModel;   // .xlsx

namespace SchemaStudio
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance); // ExcelDataReader 需要
            ApplicationConfiguration.Initialize();
            Application.Run(new MainForm());
        }
    }

    public class MainForm : Form
    {
        private readonly SplitContainer split;
        private readonly Panel topBar;
        private readonly Button btnOpenDir, btnSave, btnCheck, btnNewTable, btnPropLib;
        private readonly Label lblDir;
        private readonly TreeView tree;
        private readonly DataGridView grid;
        private readonly BindingList<ColumnRow> rows = new();
        private readonly ToolTip tips = new();

        private string? currentDir;
        private NodeTag? currentTag;               // 兼容文件/Sheet 两级
        private TableSchema? currentSchema;        // 当前编辑的「单表/Sheet」结构
        private WorkbookSchema? currentWorkbook;   // 当前 Excel 文件对应的工作簿 JSON（保存时写回）
        private PropertyRegistry registry = new();

        private static readonly string[] FileExts = new[] { ".csv", ".xlsx", ".xls" };
        private static readonly string[] AllowedTypes = new[]
        {
            "int","long","float","double","bool","string","int[]","List<int>"
        };

        private string RegistryPath => currentDir == null ? "" : Path.Combine(currentDir!, "schema.registry.json");
        private string UiStatePath  => currentDir == null ? "" : Path.Combine(currentDir!, "schema.ui.json");

        public MainForm()
        {
            Text = "SchemaStudio — JSON 可视化编辑器";
            Width = 1280; Height = 760;

            // 更稳的缩放模式（高分屏不容易出错）
            AutoScaleMode = AutoScaleMode.Dpi;

            // ====== 顶部栏控件 ======
            btnOpenDir = new Button { Text = "打开目录",   Left = 10,  Width = 90, Top = 8 };
            btnSave    = new Button { Text = "保存JSON",   Left = 110, Width = 90, Top = 8 };
            btnCheck   = new Button { Text = "检查",       Left = 210, Width = 70, Top = 8 };
            btnNewTable= new Button { Text = "新建表",     Left = 290, Width = 80, Top = 8 };
            btnPropLib = new Button { Text = "自定义属性", Left = 380, Width = 110, Top = 8 };
            lblDir     = new Label  { Text = "未选择目录", AutoSize = true, Left = 510, Top = 12 };

            // 事件
            btnOpenDir.Click += (_, __) => ChooseDir();
            btnSave.Click    += (_, __) => SaveSchema();
            btnCheck.Click   += (_, __) => ValidateCurrent();
            btnNewTable.Click+= (_, __) => CreateNewTable();
            btnPropLib.Click += (_, __) => OpenPropertyLibrary();

            tips.SetToolTip(btnCheck,   "校验当前表数据与配置的数值类型是否匹配");
            tips.SetToolTip(btnNewTable,"新建 CSV / XLSX / XLS（Excel 可创建多个工作表）");
            tips.SetToolTip(btnPropLib, "编辑列自定义属性库(根目录 schema.registry.json，团队共享)");

            // 顶部栏：使用 AutoSize，避免被裁切
            topBar = new Panel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(8, 6, 8, 6)
            };
            topBar.Controls.AddRange(new Control[] { btnOpenDir, btnSave, btnCheck, btnNewTable, btnPropLib, lblDir });

            // ====== 中间分栏 ======
            split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                SplitterDistance = 380
            };

            // 左树（懒加载 Excel Sheets）
            tree = new TreeView { Dock = DockStyle.Fill, HideSelection = false, ShowLines = true };
            tree.BeforeExpand += (_, e) =>
            {
                if (e.Node?.Tag is NodeTag t && t.IsExcelFile && e.Node.Nodes.Count == 1 && e.Node.Nodes[0].Tag == null)
                    AddSheetChildren(e.Node, t.Path);
            };
            tree.AfterSelect += (_, e) => TryLoadNode(e.Node);
            tree.NodeMouseDoubleClick += (_, e) => TryLoadNode(e.Node);
            split.Panel1.Controls.Add(tree);

            // 右表格
            grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                AllowUserToAddRows = false,
                RowHeadersVisible = false,
                // 让表头高度自适应，避免顶部看起来像被“吃掉”
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
            };
            BuildGridColumns();
            grid.DataSource = rows;
            grid.CellContentClick += Grid_CellContentClick;
            grid.ColumnWidthChanged += (_, __) => SaveUiState();
            split.SplitterMoved += (_, __) => SaveUiState();
            split.Panel2.Controls.Add(grid);

            // ====== 根布局：TableLayoutPanel 两行 ======
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));           // 第1行：topBar 自适应
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));      // 第2行：split 占满
            root.Controls.Add(topBar, 0, 0);
            root.Controls.Add(split, 0, 1);

            // 把根布局加到窗体（不要直接把 topBar/split 加到窗体）
            Controls.Add(root);

            this.FormClosing += (_, __) => SaveUiState();
        }

        private void BuildGridColumns()
        {
            grid.Columns.Clear();
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name="col_name", HeaderText="列名", DataPropertyName=nameof(ColumnRow.ColName), ReadOnly=true, Width=260 });
            grid.Columns.Add(new DataGridViewComboBoxColumn { Name="value_type", HeaderText="数值类型", DataPropertyName=nameof(ColumnRow.ValueType), Width=140, DataSource=AllowedTypes });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name="pk_order", HeaderText="主键序列", DataPropertyName=nameof(ColumnRow.PkOrderText), Width=150 });
            grid.Columns.Add(new DataGridViewButtonColumn { Name="prop_btn", HeaderText="列属性", Text="属性…", UseColumnTextForButtonValue=true, Width=80 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name="prop_summary", HeaderText="属性摘要", DataPropertyName=nameof(ColumnRow.PropsSummary), ReadOnly=true, AutoSizeMode=DataGridViewAutoSizeColumnMode.Fill });
        }

        // ======= 顶部动作 =======
        private void ChooseDir()
        {
            using var fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                currentDir = fbd.SelectedPath;
                lblDir.Text = currentDir;
                LoadRegistry();
                BuildDirectoryTree();
                ApplyUiState();
            }
        }

        private void LoadRegistry()
        {
            try
            {
                registry = File.Exists(RegistryPath)
                    ? JsonConvert.DeserializeObject<PropertyRegistry>(File.ReadAllText(RegistryPath)) ?? new PropertyRegistry()
                    : new PropertyRegistry();

                EnsureProp("cs_type", "终端类型", scope: "column", kind: "enum", values: new[] { "ALL","C","S" }, defVal: "ALL");
                EnsureProp("translate_mode", "翻译类型", scope: "column", kind: "enum", values: new[] { "none","auto","manual" }, defVal: "none");
                SaveRegistry();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"读取属性库失败：{ex.Message}");
                registry = new PropertyRegistry();
            }
        }
        private void EnsureProp(string key, string title, string scope, string kind, IEnumerable<string> values, string defVal)
        {
            if (!registry.Properties.Any(p => p.Key.Equals(key, StringComparison.OrdinalIgnoreCase)))
            {
                registry.Properties.Add(new PropertyDef { Key = key, Title = title, Scope = scope, Kind = kind, EnumValues = values.ToList(), Default = defVal });
            }
        }
        private void SaveRegistry()
        {
            if (string.IsNullOrEmpty(RegistryPath)) return;
            var json = JsonConvert.SerializeObject(registry, Formatting.Indented,
                new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
            File.WriteAllText(RegistryPath, json, new UTF8Encoding(false));
        }

        private void OpenPropertyLibrary()
        {
            if (currentDir == null) { MessageBox.Show("请先打开目录"); return; }
            using var dlg = new PropertyLibraryForm(registry);
            if (dlg.ShowDialog(this) == DialogResult.OK)
            {
                registry = dlg.Result;
                SaveRegistry();
            }
        }

        // ======= 目录树（懒加载 Sheets）=======
        private void BuildDirectoryTree()
        {
            tree.BeginUpdate();
            try
            {
                tree.Nodes.Clear();
                if (currentDir == null || !Directory.Exists(currentDir)) return;

                var root = new TreeNode(Path.GetFileName(currentDir.TrimEnd(Path.DirectorySeparatorChar))) { Tag = new NodeTag(currentDir, null) };
                tree.Nodes.Add(root);
                AddDirNodes(root, currentDir);
                root.Expand();
            }
            finally { tree.EndUpdate(); }
        }

        private void AddDirNodes(TreeNode parent, string dir)
        {
            foreach (var sub in Directory.GetDirectories(dir).OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
            {
                var node = new TreeNode(Path.GetFileName(sub)) { Tag = new NodeTag(sub, null) };
                parent.Nodes.Add(node);
                AddDirNodes(node, sub);
            }
            foreach (var file in Directory.GetFiles(dir)
                         .Where(p => FileExts.Contains(Path.GetExtension(p).ToLowerInvariant()))
                         .Where(p => !Path.GetFileName(p).StartsWith("~$"))
                         .OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
            {
                var ext = Path.GetExtension(file).ToLowerInvariant();
                var node = new TreeNode(Path.GetFileName(file)) { Tag = new NodeTag(file, null) };
                parent.Nodes.Add(node);
                if (ext is ".xlsx" or ".xls")
                    node.Nodes.Add(new TreeNode("(展开查看工作表)")); // 懒加载
            }
        }

        private void AddSheetChildren(TreeNode fileNode, string excelPath)
        {
            fileNode.Nodes.Clear();
            try
            {
                using var fs = File.OpenRead(excelPath);
                using var r = ExcelReaderFactory.CreateReader(fs);
                var ds = r.AsDataSet();
                foreach (DataTable t in ds.Tables)
                    fileNode.Nodes.Add(new TreeNode(t.TableName) { Tag = new NodeTag(excelPath, t.TableName) });
            }
            catch (Exception ex)
            {
                fileNode.Nodes.Add(new TreeNode($"(读取失败: {ex.Message})"));
            }
        }

        private void TryLoadNode(TreeNode? node)
        {
            if (node?.Tag is not NodeTag tag) return;
            currentTag = tag;
            currentSchema = null;
            currentWorkbook = null;

            if (tag.IsExcelFile)
            {
                rows.Clear(); // 文件层不显示数据；展开看 Sheet
                Text = $"SchemaStudio — {Path.GetFileName(tag.Path)}";
                return;
            }

            if (tag.IsExcelSheet) LoadExcelSheet(tag.Path, tag.Sheet!);
            else if (tag.IsCsv)   LoadCsv(tag.Path);
        }

        // ======= 加载 CSV / Excel+sheet =======
        private void LoadCsv(string path)
        {
            var tableName = Path.GetFileName(path);
            var jsonPath  = GetSchemaJsonPath(path);

            var colNames = ReadCsvHeader(path);
            if (colNames.Count == 0) { MessageBox.Show("未能读取到列名。请确认第1行是字段名。"); return; }

            var schema = File.Exists(jsonPath)
                ? JsonConvert.DeserializeObject<TableSchema>(File.ReadAllText(jsonPath)) ?? new TableSchema { TableName = tableName }
                : new TableSchema { TableName = tableName };
            var cols = MapColumns(colNames, schema.Columns);
            currentSchema = new TableSchema { TableName = tableName, Columns = cols };
            rows.Clear(); foreach (var c in cols) rows.Add(ColumnRow.FromSchema(c));
            Text = $"SchemaStudio — {tableName}";
        }

        private void LoadExcelSheet(string path, string sheetName)
        {
            var tableName = $"{Path.GetFileName(path)}::{sheetName}";
            var jsonPath  = GetSchemaJsonPath(path);

            var colNames = ReadXlsxHeader(path, sheetName);
            if (colNames.Count == 0) { MessageBox.Show("未能读取到列名。请确认第1行是字段名。"); return; }

            currentWorkbook = File.Exists(jsonPath)
                ? JsonConvert.DeserializeObject<WorkbookSchema>(File.ReadAllText(jsonPath))
                : null;

            // 兼容历史：若是单表 JSON 也能读
            if (currentWorkbook == null && File.Exists(jsonPath))
            {
                var legacy = JsonConvert.DeserializeObject<TableSchema>(File.ReadAllText(jsonPath));
                if (legacy?.Columns != null)
                {
                    currentWorkbook = new WorkbookSchema
                    {
                        WorkbookName = Path.GetFileName(path),
                        Sheets = new List<SheetSchema> { new SheetSchema { SheetName = sheetName, Columns = legacy.Columns } }
                    };
                }
            }
            currentWorkbook ??= new WorkbookSchema { WorkbookName = Path.GetFileName(path), Sheets = new List<SheetSchema>() };

            var entry = currentWorkbook.Sheets.FirstOrDefault(s => s.SheetName.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
            var cols  = MapColumns(colNames, entry?.Columns);
            currentSchema = new TableSchema { TableName = tableName, Columns = cols };
            rows.Clear(); foreach (var c in cols) rows.Add(ColumnRow.FromSchema(c));
            Text = $"SchemaStudio — {tableName}";
        }

        private List<ColumnSchema> MapColumns(List<string> colNames, List<ColumnSchema>? existing)
        {
            var map = existing?.ToDictionary(c => c.ColName ?? "", StringComparer.OrdinalIgnoreCase) ?? new();
            var list = new List<ColumnSchema>();
            for (int i = 0; i < colNames.Count; i++)
            {
                var n = colNames[i];
                if (map.TryGetValue(n, out var c)) { c.ColSort = i; list.Add(c); }
                else list.Add(new ColumnSchema { ColName = n, ValueType = "string", PkOrder = -1, PrimaryKey = false, ColSort = i, ColProperties = new() });
            }
            return list;
        }

        // ======= 保存 =======
        private void SaveSchema()
        {
            if (currentTag == null || currentSchema == null) return;

            // 同步 UI -> 模型
            foreach (var r in rows)
            {
                var c = currentSchema.Columns!.First(x => x.ColName == r.ColName);
                c.ValueType = r.ValueType;
                c.PkOrder   = r.PkOrder;
                c.PrimaryKey = r.PkOrder >= 0;
            }

            var jsonPath = GetSchemaJsonPath(currentTag.Path);

            if (currentTag.IsCsv)
            {
                var json = JsonConvert.SerializeObject(currentSchema, Formatting.Indented,
                    new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
                File.WriteAllText(jsonPath, json, new UTF8Encoding(false));
            }
            else if (currentTag.IsExcelSheet)
            {
                currentWorkbook ??= new WorkbookSchema { WorkbookName = Path.GetFileName(currentTag.Path), Sheets = new() };
                var entry = currentWorkbook.Sheets.FirstOrDefault(s => s.SheetName.Equals(currentTag.Sheet!, StringComparison.OrdinalIgnoreCase));
                if (entry == null) { entry = new SheetSchema { SheetName = currentTag.Sheet!, Columns = new() }; currentWorkbook.Sheets.Add(entry); }
                entry.Columns = currentSchema.Columns!;
                var json = JsonConvert.SerializeObject(currentWorkbook, Formatting.Indented,
                    new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
                File.WriteAllText(jsonPath, json, new UTF8Encoding(false));
            }
            else if (currentTag.IsExcelFile)
            {
                MessageBox.Show("请选择具体工作表后再保存。");
                return;
            }

            MessageBox.Show($"已保存：{jsonPath}", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // ======= 校验 =======
        private void ValidateCurrent()
        {
            if (currentTag == null || currentSchema == null)
            { MessageBox.Show("请先选择 CSV 或 Excel 的某个工作表。"); return; }

            var headers = currentTag.IsExcelSheet ? ReadXlsxHeader(currentTag.Path, currentTag.Sheet!) : ReadCsvHeader(currentTag.Path);
            if (headers.Count == 0) { MessageBox.Show("未能读取到列名。"); return; }
            var data = currentTag.IsExcelSheet ? ReadXlsxRows(currentTag.Path, currentTag.Sheet!, headers.Count) : ReadCsvRows(currentTag.Path, headers.Count);

            var headerIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < headers.Count; i++) headerIndex[headers[i]] = i;

            var errors = new List<ValidationError>();
            for (int rowIdx = 0; rowIdx < data.Count; rowIdx++)
            {
                var row = data[rowIdx];
                foreach (var col in currentSchema.Columns!)
                {
                    if (!headerIndex.TryGetValue(col.ColName ?? "", out var ci)) continue;
                    var val = ci < row.Length ? row[ci] ?? "" : "";
                    if (!IsValid(val, col.ValueType ?? "string", out var reason))
                    {
                        errors.Add(new ValidationError
                        {
                            Row = rowIdx + 2, Column = col.ColName ?? "", Value = val,
                            Expected = col.ValueType ?? "string", Message = reason
                        });
                    }
                }
            }
            using var dlg = new ValidationResultForm(currentSchema.TableName ?? "", errors);
            dlg.ShowDialog(this);
        }

        private static bool IsValid(string value, string type, out string reason)
        {
            reason = ""; if (string.IsNullOrWhiteSpace(value)) return true;
            var t = (type ?? "string").Trim(); var ic = CultureInfo.InvariantCulture;

            bool ParseBool(string s)
            {
                s = s.Trim().ToLowerInvariant();
                if (bool.TryParse(s, out _)) return true;
                return s is "1" or "0" or "y" or "n" or "yes" or "no" or "是" or "否" or "真" or "假";
            }

            switch (t)
            {
                case "int":    if (!int.TryParse(value, NumberStyles.Integer, ic, out _))   { reason = "不是有效的 int"; return false; } return true;
                case "long":   if (!long.TryParse(value, NumberStyles.Integer, ic, out _))  { reason = "不是有效的 long"; return false; } return true;
                case "float":  if (!float.TryParse(value, NumberStyles.Float, ic, out _))   { reason = "不是有效的 float"; return false; } return true;
                case "double": if (!double.TryParse(value, NumberStyles.Float, ic, out _))  { reason = "不是有效的 double"; return false; } return true;
                case "bool":   if (!ParseBool(value)) { reason = "不是有效的 bool(true/false/1/0/是/否)"; return false; } return true;
                case "string": return true;
                case "int[]":
                case "List<int>":
                    return TryParseIntArray(value, out _, out reason);
                default:       return true;
            }
        }
        private static bool TryParseIntArray(string s, out int[] result, out string reason)
        {
            result = Array.Empty<int>(); reason = ""; if (string.IsNullOrWhiteSpace(s)) return true; s = s.Trim();
            if (s.StartsWith("[") && s.EndsWith("]"))
            {
                try { result = (JsonConvert.DeserializeObject<List<int>>(s) ?? new()).ToArray(); return true; }
                catch { reason = "不是有效的 JSON 数组（示例：[1,2,3]）"; return false; }
            }
            var parts = s.Split(new[] { ',', '，', ';', '|', ' ' }, StringSplitOptions.RemoveEmptyEntries);
            var list = new List<int>();
            foreach (var p in parts) { if (!int.TryParse(p.Trim(), out var v)) { reason = "不是有效的整型数组（建议使用 [1,2,3]）"; return false; } list.Add(v); }
            result = list.ToArray(); return true;
        }

        // ======= 新建表（支持多 Sheet；NPOI 写入）=======
        private void CreateNewTable()
        {
            if (currentDir == null) { MessageBox.Show("请先打开根目录"); return; }
            var dirs = GetAllDirs(currentDir).ToList();
            using var dlg = new NewTableForm(currentDir, dirs, registry, AllowedTypes);
            if (dlg.ShowDialog(this) != DialogResult.OK) return;

            var destDir    = dlg.ResultDir!;
            var fileBase   = dlg.ResultFileName!;
            var fmt        = dlg.ResultFormat;               // csv/xlsx/xls
            var sheetNames = dlg.ResultSheets;               // 仅 Excel

            string path = Path.Combine(destDir, fileBase + (fmt == "xlsx" ? ".xlsx" : fmt == "xls" ? ".xls" : ".csv"));
            if (File.Exists(path)) { MessageBox.Show("同名文件已存在，已取消。"); return; }
            Directory.CreateDirectory(destDir);

            var headers = dlg.Columns.Select(c => c.ColName).ToArray();
            bool wrote = false; string warn = "";

            try
            {
                if (fmt == "xlsx") wrote = TryWriteXlsxMulti(path, sheetNames, headers, out warn);
                else if (fmt == "xls") wrote = TryWriteXlsMulti(path, sheetNames, headers, out warn);
            }
            catch (Exception ex) { warn = $"写入 Excel 失败：{ex.Message}"; wrote = false; }

            if (!wrote) // 回退 CSV
            {
                path = Path.Combine(destDir, fileBase + ".csv");
                WriteCsvHeader(path, headers);
                if (!string.IsNullOrEmpty(warn)) MessageBox.Show(warn + "\n已回退为 CSV。");
            }

            // 生成 schema JSON
            var jsonPath = GetSchemaJsonPath(path);
            if (fmt == "csv" || !wrote)
            {
                var schema = new TableSchema
                {
                    TableName = Path.GetFileName(path),
                    Columns = dlg.Columns.Select((c, idx) => new ColumnSchema
                    {
                        ColName = c.ColName, ValueType = c.ValueType, PkOrder = c.PkOrder,
                        PrimaryKey = c.PkOrder >= 0, ColSort = idx, ColProperties = c.ColProperties?.ToList() ?? new()
                    }).ToList()
                };
                var json = JsonConvert.SerializeObject(schema, Formatting.Indented,
                    new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
                File.WriteAllText(jsonPath, json, new UTF8Encoding(false));
            }
            else
            {
                var wb = new WorkbookSchema
                {
                    WorkbookName = Path.GetFileName(path),
                    Sheets = sheetNames.Select(sn => new SheetSchema
                    {
                        SheetName = sn,
                        Columns = dlg.Columns.Select((c, idx) => new ColumnSchema
                        {
                            ColName = c.ColName, ValueType = c.ValueType, PkOrder = c.PkOrder,
                            PrimaryKey = c.PkOrder >= 0, ColSort = idx, ColProperties = c.ColProperties?.ToList() ?? new()
                        }).ToList()
                    }).ToList()
                };
                var json = JsonConvert.SerializeObject(wb, Formatting.Indented,
                    new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
                File.WriteAllText(jsonPath, json, new UTF8Encoding(false));
            }

            BuildDirectoryTree();

            // 定位：CSV 直接选中；Excel 选中文件节点（展开可见所有 Sheet）
            SelectPathInTree(path);
            MessageBox.Show($"已创建：\n{path}\n{path}.schema.json");
        }

        // —— NPOI 写 Excel 首行（多 Sheet）——
        private static bool TryWriteXlsxMulti(string path, IEnumerable<string> sheets, IEnumerable<string> headers, out string warn)
        {
            warn = "";
            using var wb = new XSSFWorkbook();
            foreach (var sn in sheets)
            {
                var ws = wb.CreateSheet(sn);
                var r0 = ws.CreateRow(0);
                int c = 0; foreach (var h in headers) r0.CreateCell(c++).SetCellValue(h);
            }
            using var fs = new FileStream(path, FileMode.Create, FileAccess.Write);
            wb.Write(fs);
            return true;
        }
        private static bool TryWriteXlsMulti(string path, IEnumerable<string> sheets, IEnumerable<string> headers, out string warn)
        {
            warn = "";
            using var wb = new HSSFWorkbook();
            foreach (var sn in sheets)
            {
                var ws = wb.CreateSheet(sn);
                var r0 = ws.CreateRow(0);
                int c = 0; foreach (var h in headers) r0.CreateCell(c++).SetCellValue(h);
            }
            using var fs = new FileStream(path, FileMode.Create, FileAccess.Write);
            wb.Write(fs);
            return true;
        }

        // —— 写 CSV 首行 —— 
        private static void WriteCsvHeader(string path, IEnumerable<string> headers)
        {
            using var sw = new StreamWriter(path, false, new UTF8Encoding(true));
            static string Esc(string s) =>
                (s.Contains('"') || s.Contains(',') || s.Contains('\n') || s.Contains('\r'))
                    ? "\"" + s.Replace("\"", "\"\"") + "\"" : s;
            sw.WriteLine(string.Join(",", headers.Select(Esc)));
        }

        // —— 读取 CSV/Excel —— 
        private static List<string> ReadCsvHeader(string path)
        {
            using var reader = new StreamReader(path, DetectEncoding(path));
            using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
            if (!csv.Read()) return new();
            var header = new List<string>();
            for (int i = 0; csv.TryGetField<string>(i, out var val); i++) header.Add((val ?? $"col_{i}").Trim());
            return header;
        }
        private static List<string[]> ReadCsvRows(string path, int expectedCols)
        {
            var list = new List<string[]>();
            using var reader = new StreamReader(path, DetectEncoding(path));
            using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
            if (!csv.Read()) return list;
            while (csv.Read())
            {
                var row = new string[expectedCols];
                for (int i = 0; i < expectedCols; i++) { csv.TryGetField<string>(i, out var v); row[i] = v?.Trim() ?? ""; }
                list.Add(row);
            }
            return list;
        }
        private static List<string> ReadXlsxHeader(string path, string sheet)
        {
            using var fs = File.OpenRead(path);
            using var r  = ExcelReaderFactory.CreateReader(fs);
            var ds = r.AsDataSet();
            var t  = ds.Tables.Cast<DataTable>().FirstOrDefault(x => string.Equals(x.TableName, sheet, StringComparison.OrdinalIgnoreCase)) ?? ds.Tables[0];
            if (t.Rows.Count == 0) return new();
            var row = t.Rows[0];
            var list = new List<string>();
            for (int i = 0; i < t.Columns.Count; i++) list.Add((row[i]?.ToString() ?? $"col_{i}").Trim());
            return list;
        }
        private static List<string[]> ReadXlsxRows(string path, string sheet, int expectedCols)
        {
            var list = new List<string[]>();
            using var fs = File.OpenRead(path);
            using var r  = ExcelReaderFactory.CreateReader(fs);
            var ds = r.AsDataSet();
            var t  = ds.Tables.Cast<DataTable>().FirstOrDefault(x => string.Equals(x.TableName, sheet, StringComparison.OrdinalIgnoreCase)) ?? ds.Tables[0];
            if (t.Rows.Count <= 1) return list;
            for (int rr = 1; rr < t.Rows.Count; rr++)
            {
                var dr = t.Rows[rr];
                var row = new string[expectedCols];
                for (int c = 0; c < expectedCols; c++) row[c] = dr[c]?.ToString()?.Trim() ?? "";
                list.Add(row);
            }
            return list;
        }

        // —— 其它 —— 
        private static string GetSchemaJsonPath(string tablePath)
            => Path.Combine(Path.GetDirectoryName(tablePath)!, Path.GetFileName(tablePath) + ".schema.json");
        private IEnumerable<string> GetAllDirs(string root)
        {
            yield return root;
            foreach (var d in Directory.GetDirectories(root))
                foreach (var x in GetAllDirs(d)) yield return x;
        }
        private void ApplyUiState()
        {
            if (string.IsNullOrEmpty(UiStatePath) || !File.Exists(UiStatePath)) return;
            try
            {
                var state = JsonConvert.DeserializeObject<UiState>(File.ReadAllText(UiStatePath));
                if (state == null) return;
                if (state.SplitterDistance > 0 && state.SplitterDistance < split.Width - 80)
                    split.SplitterDistance = state.SplitterDistance;
                foreach (DataGridViewColumn c in grid.Columns)
                    if (state.ColWidths.TryGetValue(c.Name, out var w) && w > 0) c.Width = w;
            }
            catch { }
        }
        private void SaveUiState()
        {
            if (string.IsNullOrEmpty(UiStatePath)) return;
            var st = new UiState { SplitterDistance = split.SplitterDistance, ColWidths = new Dictionary<string, int>() };
            foreach (DataGridViewColumn c in grid.Columns) st.ColWidths[c.Name] = c.Width;
            File.WriteAllText(UiStatePath, JsonConvert.SerializeObject(st, Formatting.Indented));
        }
        private static Encoding DetectEncoding(string filePath)
        {
            var bytes = File.ReadAllBytes(filePath);
            try { Encoding.UTF8.GetString(bytes); return new UTF8Encoding(false); }
            catch { return Encoding.UTF8; }
        }

        private void Grid_CellContentClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            var btnColIndex = grid.Columns["prop_btn"]!.Index;
            if (e.ColumnIndex == btnColIndex && grid.Rows[e.RowIndex].DataBoundItem is ColumnRow row)
            {
                if (currentSchema == null) return;
                var col  = currentSchema.Columns!.First(c => c.ColName == row.ColName);
                var defs = registry.Properties.Where(d => d.Scope == "column" || d.Scope == "both").ToList();

                using var dlg = new AssignPropertyForm(defs, col.ColProperties ?? new List<PropertyKV>());
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    col.ColProperties = dlg.Result;
                    row.PropsSummary  = string.Join("; ", (col.ColProperties ?? new()).Select(p => $"{p.PropertyName}={p.PropertyValue}"));
                    grid.InvalidateRow(e.RowIndex);
                }
            }
        }

        private void SelectPathInTree(string fullPath)
        {
            foreach (TreeNode root in tree.Nodes)
            {
                var node = FindNode(root, fullPath);
                if (node != null)
                {
                    tree.SelectedNode = node; node.EnsureVisible(); break;
                }
            }
            TreeNode? FindNode(TreeNode n, string path)
            {
                if (n.Tag is NodeTag t && t.Path == path && t.Sheet == null) return n;
                foreach (TreeNode c in n.Nodes) { var got = FindNode(c, path); if (got != null) return got; }
                return null;
            }
        }
    }

    // —— 树节点 Tag —— 
    public record NodeTag(string Path, string? Sheet)
    {
        public bool IsExcelFile  => (Path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || Path.EndsWith(".xls", StringComparison.OrdinalIgnoreCase)) && Sheet == null;
        public bool IsExcelSheet => Sheet != null && (Path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || Path.EndsWith(".xls", StringComparison.OrdinalIgnoreCase));
        public bool IsCsv        => Path.EndsWith(".csv", StringComparison.OrdinalIgnoreCase);
        public override string ToString() => Sheet == null ? Path : $"{Path}::{Sheet}";
    }

    // —— 新建表对话框（多 Sheet，NPOI 写入） —— 
    public class NewTableForm : Form
    {
        private readonly ComboBox cmbDir, cmbFormat;
        private readonly TextBox txtName;
        private readonly DataGridView gridCols, gridSheets;
        private readonly Button btnAddCol, btnDelCol, btnAddSheet, btnDelSheet, btnOk, btnCancel;

        private readonly BindingList<NewColRow> cols = new();
        private readonly BindingList<SheetNameRow> sheets = new();

        private readonly PropertyRegistry registry;
        private readonly string root;
        private readonly string[] allowed;

        public string? ResultDir { get; private set; }
        public string? ResultFileName { get; private set; }
        public string ResultFormat { get; private set; } = "csv";                 // csv/xlsx/xls
        public List<string> ResultSheets => sheets.Select(s => s.Name).Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
        public IEnumerable<NewColRow> Columns => cols;

        public NewTableForm(string root, List<string> allDirs, PropertyRegistry reg, string[] allowedTypes)
        {
            this.root = root; registry = reg; allowed = allowedTypes;

            Text = "新建表"; Width = 980; Height = 640; AutoScaleMode = AutoScaleMode.Font;

            var lblDir = new Label { Left = 12, Top = 14, Text = "保存到目录：", AutoSize = true };
            cmbDir = new ComboBox { Left = 90, Top = 10, Width = 680, DropDownStyle = ComboBoxStyle.DropDownList, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            cmbDir.Items.AddRange(allDirs.Select(d => d.Replace(root, "").TrimStart(Path.DirectorySeparatorChar)).Select(s => string.IsNullOrEmpty(s) ? "(根目录)" : s).ToArray());
            cmbDir.SelectedIndex = 0;

            var lblName = new Label { Left = 12, Top = 45, Text = "文件名：", AutoSize = true };
            txtName = new TextBox { Left = 90, Top = 41, Width = 260, Text = "new_table" };

            var lblFmt = new Label { Left = 360, Top = 45, Text = "格式：", AutoSize = true };
            cmbFormat = new ComboBox { Left = 400, Top = 41, Width = 100, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbFormat.Items.AddRange(new[] { "csv", "xlsx", "xls" });
            cmbFormat.SelectedIndex = 0;
            cmbFormat.SelectedIndexChanged += (_, __) =>
            {
                ResultFormat = cmbFormat.SelectedItem!.ToString()!;
                var isExcel = ResultFormat is "xlsx" or "xls";
                gridSheets.Enabled = btnAddSheet.Enabled = btnDelSheet.Enabled = isExcel;
                gridSheets.Visible = btnAddSheet.Visible = btnDelSheet.Visible = isExcel;
            };

            // 列定义
            gridCols = new DataGridView
            {
                Left = 12, Top = 80, Width = 940, Height = 380,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                AutoGenerateColumns = false, AllowUserToAddRows = false
            };
            gridCols.DataSource = cols;
            gridCols.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "列名", DataPropertyName = nameof(NewColRow.ColName), Width = 240 });
            gridCols.Columns.Add(new DataGridViewComboBoxColumn { HeaderText = "数值类型", DataPropertyName = nameof(NewColRow.ValueType), Width = 140, DataSource = allowed });
            gridCols.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "主键序列", DataPropertyName = nameof(NewColRow.PkOrderText), Width = 120 });
            var btn = new DataGridViewButtonColumn { HeaderText = "列属性", Text = "属性…", UseColumnTextForButtonValue = true, Width = 80 };
            gridCols.Columns.Add(btn);
            gridCols.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "属性摘要", DataPropertyName = nameof(NewColRow.PropsSummary), AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, ReadOnly = true });
            gridCols.CellContentClick += (s, e) =>
            {
                if (e.RowIndex < 0) return;
                if (gridCols.Columns[e.ColumnIndex] is DataGridViewButtonColumn && gridCols.Rows[e.RowIndex].DataBoundItem is NewColRow row)
                {
                    var defs = registry.Properties.Where(d => d.Scope == "column" || d.Scope == "both").ToList();
                    using var dlg = new AssignPropertyForm(defs, row.ColProperties ?? new List<PropertyKV>());
                    if (dlg.ShowDialog(this) == DialogResult.OK)
                    {
                        row.ColProperties = dlg.Result;
                        row.PropsSummary = string.Join("; ", row.ColProperties.Select(p => $"{p.PropertyName}={p.PropertyValue}"));
                        gridCols.InvalidateRow(e.RowIndex);
                    }
                }
            };

            btnAddCol = new Button { Text = "添加列", Left = 12, Top = 470, Width = 80, Anchor = AnchorStyles.Bottom | AnchorStyles.Left };
            btnDelCol = new Button { Text = "删除选中", Left = 100, Top = 470, Width = 90, Anchor = AnchorStyles.Bottom | AnchorStyles.Left };
            btnAddCol.Click += (_, __) => cols.Add(new NewColRow { ColName = $"col_{cols.Count + 1}", ValueType = "string", PkOrder = -1 });
            btnDelCol.Click += (_, __) => { foreach (DataGridViewRow r in gridCols.SelectedRows) if (r.DataBoundItem is NewColRow x) cols.Remove(x); };

            // Sheet 列表（仅 Excel 启用）
            gridSheets = new DataGridView
            {
                Left = 12, Top = 510, Width = 480, Height = 80,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left,
                AutoGenerateColumns = false, AllowUserToAddRows = false,
                Visible = false, Enabled = false
            };
            gridSheets.DataSource = sheets;
            gridSheets.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Sheet 名", DataPropertyName = nameof(SheetNameRow.Name), Width = 360 });
            btnAddSheet = new Button { Text = "添加Sheet", Left = 510, Top = 516, Width = 90, Anchor = AnchorStyles.Bottom | AnchorStyles.Left, Visible = false, Enabled = false };
            btnDelSheet = new Button { Text = "删除选中", Left = 610, Top = 516, Width = 90, Anchor = AnchorStyles.Bottom | AnchorStyles.Left, Visible = false, Enabled = false };
            btnAddSheet.Click += (_, __) => sheets.Add(new SheetNameRow { Name = $"Sheet{sheets.Count + 1}" });
            btnDelSheet.Click += (_, __) => { foreach (DataGridViewRow r in gridSheets.SelectedRows) if (r.DataBoundItem is SheetNameRow x) sheets.Remove(x); };

            btnOk     = new Button { Text = "确定", Left = 780, Top = 560, Width = 80, Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            btnCancel = new Button { Text = "取消", Left = 872, Top = 560, Width = 80, Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            btnOk.Click += (_, __) => OnOk(allDirs);
            btnCancel.Click += (_, __) => { DialogResult = DialogResult.Cancel; Close(); };

            Controls.AddRange(new Control[] { lblDir, cmbDir, lblName, txtName, lblFmt, cmbFormat, gridCols, btnAddCol, btnDelCol, gridSheets, btnAddSheet, btnDelSheet, btnOk, btnCancel });

            // 默认两列 + 一个 sheet
            cols.Add(new NewColRow { ColName = "id",   ValueType = "int",    PkOrder = 1 });
            cols.Add(new NewColRow { ColName = "desc", ValueType = "string", PkOrder = -1 });
            sheets.Add(new SheetNameRow { Name = "Sheet1" });
        }

        private void OnOk(List<string> allDirs)
        {
            var name = (txtName.Text ?? "").Trim();
            if (string.IsNullOrEmpty(name)) { MessageBox.Show("文件名不能为空"); return; }
            if (name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) { MessageBox.Show("文件名包含非法字符"); return; }
            var rel = cmbDir.SelectedItem!.ToString()!;
            string destDir = rel == "(根目录)" ? root : Path.Combine(root, rel.TrimStart(Path.DirectorySeparatorChar));
            if (!allDirs.Contains(destDir)) { MessageBox.Show("目标目录无效"); return; }
            if (cols.Count == 0) { MessageBox.Show("请至少添加一列"); return; }

            ResultFormat = cmbFormat.SelectedItem!.ToString()!;
            if (ResultFormat is "xlsx" or "xls")
            {
                if (sheets.Count == 0) { MessageBox.Show("请至少添加一个工作表"); return; }
                var valid = sheets.All(s => !string.IsNullOrWhiteSpace(s.Name));
                if (!valid) { MessageBox.Show("存在空的工作表名"); return; }
                var distinct = sheets.Select(s => s.Name).Distinct(StringComparer.OrdinalIgnoreCase).Count() == sheets.Count;
                if (!distinct) { MessageBox.Show("工作表名重复"); return; }
            }

            ResultDir = destDir;
            ResultFileName = name;
            DialogResult = DialogResult.OK;
            Close();
        }
    }

    public class NewColRow : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;
        public string ColName { get; set; } = "";
        public string ValueType { get; set; } = "string";
        public int PkOrder { get; set; } = -1;
        public List<PropertyKV>? ColProperties { get; set; } = new();
        public string PropsSummary { get; set; } = "";
        public string PkOrderText
        {
            get => PkOrder < 0 ? string.Empty : PkOrder.ToString();
            set { if (string.IsNullOrWhiteSpace(value)) PkOrder = -1; else if (int.TryParse(value.Trim(), out var v)) PkOrder = v; else PkOrder = -1;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(PkOrderText))); }
        }
    }
    public class SheetNameRow { public string Name { get; set; } = "Sheet1"; }

    // —— 绑定模型/Schema —— 
    public class ColumnRow : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;
        public string ColName { get; set; } = "";
        public string ValueType { get; set; } = "string";
        public int PkOrder { get; set; } = -1;
        public string PropsSummary { get; set; } = "";
        public string PkOrderText
        {
            get => PkOrder < 0 ? string.Empty : PkOrder.ToString();
            set { if (string.IsNullOrWhiteSpace(value)) PkOrder = -1; else if (int.TryParse(value.Trim(), out var v)) PkOrder = v; else PkOrder = -1;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(PkOrderText))); }
        }
        public static ColumnRow FromSchema(ColumnSchema c) => new()
        {
            ColName = c.ColName ?? "",
            ValueType = string.IsNullOrWhiteSpace(c.ValueType) ? "string" : c.ValueType!,
            PkOrder = c.PkOrder,
            PropsSummary = c.ColProperties == null || c.ColProperties.Count == 0 ? "" :
                           string.Join("; ", c.ColProperties.Select(p => $"{p.PropertyName}={p.PropertyValue}"))
        };
    }

    public class TableSchema
    {
        [JsonProperty("table_name")] public string? TableName { get; set; }
        [JsonProperty("columns")]    public List<ColumnSchema>? Columns { get; set; }
    }
    public class WorkbookSchema
    {
        [JsonProperty("workbook_name")] public string? WorkbookName { get; set; }
        [JsonProperty("sheets")]        public List<SheetSchema> Sheets { get; set; } = new();
    }
    public class SheetSchema
    {
        [JsonProperty("sheet_name")] public string SheetName { get; set; } = "";
        [JsonProperty("columns")]    public List<ColumnSchema> Columns { get; set; } = new();
    }
    public class ColumnSchema
    {
        [JsonProperty("col_name")]       public string? ColName { get; set; }
        [JsonProperty("value_type")]     public string? ValueType { get; set; }
        [JsonProperty("pk_order")]       public int PkOrder { get; set; } = -1;
        [JsonProperty("primary_key")]    public bool PrimaryKey { get; set; }
        [JsonProperty("col_sort")]       public int ColSort { get; set; }
        [JsonProperty("col_properties")] public List<PropertyKV>? ColProperties { get; set; }
    }
    public class PropertyKV
    {
        public PropertyKV() { }
        public PropertyKV(string name, string value) { PropertyName = name; PropertyValue = value; }
        [JsonProperty("property_name")] public string? PropertyName { get; set; }
        [JsonProperty("property_value")] public string? PropertyValue { get; set; }
    }
    public class PropertyRegistry
    {
        [JsonProperty("version")]    public int Version { get; set; } = 1;
        [JsonProperty("properties")] public List<PropertyDef> Properties { get; set; } = new();
    }
    public class PropertyDef
    {
        [JsonProperty("key")]     public string Key { get; set; } = string.Empty;
        [JsonProperty("title")]   public string Title { get; set; } = string.Empty;
        [JsonProperty("scope")]   public string Scope { get; set; } = "column";
        [JsonProperty("kind")]    public string Kind { get; set; } = "enum"; // 默认枚举
        [JsonProperty("values")]  public List<string>? EnumValues { get; set; } = new();
        [JsonProperty("default")] public string? Default { get; set; }
        [JsonIgnore] public string Display => string.IsNullOrEmpty(Title) ? Key : $"{Key}（{Title}）";
        public override string ToString() => Display;
    }

    // —— 属性库窗 / 赋值窗 / 校验结果窗 —— 
    public class PropertyLibraryForm : Form
    {
        private readonly ListBox list;
        private readonly TextBox txtKey, txtTitle, txtEnum, txtDefault;
        private readonly Button btnAdd, btnDel, btnSave, btnOk, btnCancel;
        private PropertyRegistry working;
        public PropertyRegistry Result => working;

        public PropertyLibraryForm(PropertyRegistry reg)
        {
            Text = "自定义属性库"; Width = 680; Height = 480;
            working = JsonConvert.DeserializeObject<PropertyRegistry>(JsonConvert.SerializeObject(reg)) ?? new PropertyRegistry(); // 深拷贝

            list = new ListBox { Left = 10, Top = 10, Width = 240, Height = 380 };
            list.SelectedIndexChanged += (_, __) => LoadSelected();
            Controls.Add(list);

            int L = 260, W = 380, y = 10, gap = 34;
            Controls.Add(new Label { Left = L, Top = y, Text = "key", Width = 80 });
            txtKey = new TextBox { Left = L + 80, Top = y, Width = W - 80 }; y += gap; Controls.Add(txtKey);

            Controls.Add(new Label { Left = L, Top = y, Text = "标题(title)", Width = 90 });
            txtTitle = new TextBox { Left = L + 100, Top = y, Width = W - 100 }; y += gap; Controls.Add(txtTitle);

            Controls.Add(new Label { Left = L, Top = y, Text = "enum 值(逗号)", Width = 110 });
            txtEnum = new TextBox { Left = L + 120, Top = y, Width = W - 120 }; y += gap; Controls.Add(txtEnum);

            Controls.Add(new Label { Left = L, Top = y, Text = "默认值", Width = 80 });
            txtDefault = new TextBox { Left = L + 80, Top = y, Width = W - 80 }; y += gap; Controls.Add(txtDefault);

            btnAdd = new Button { Left = L, Top = y + 10, Text = "新增/清空", Width = 100 };
            btnDel = new Button { Left = L + 110, Top = y + 10, Text = "删除", Width = 80 };
            btnSave = new Button { Left = L + 200, Top = y + 10, Text = "保存条目", Width = 90 };
            btnAdd.Click += (_, __) => { txtKey.Text = txtTitle.Text = txtEnum.Text = txtDefault.Text = ""; list.ClearSelected(); };
            btnDel.Click += (_, __) => { if (list.SelectedItem is PropertyDef d) { working.Properties.RemoveAll(p => p.Key.Equals(d.Key, StringComparison.OrdinalIgnoreCase)); RefreshList(); } };
            btnSave.Click += (_, __) => SaveCurrent();
            Controls.AddRange(new Control[] { btnAdd, btnDel, btnSave });

            btnOk = new Button { Text = "确定", Left = Width - 200, Top = Height - 90, Width = 80, Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            btnCancel = new Button { Text = "取消", Left = Width - 110, Top = Height - 90, Width = 80, Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            btnOk.Click += (_, __) => { DialogResult = DialogResult.OK; Close(); };
            btnCancel.Click += (_, __) => { DialogResult = DialogResult.Cancel; Close(); };
            Controls.Add(btnOk); Controls.Add(btnCancel);

            RefreshList();
        }

        private void RefreshList()
        {
            list.Items.Clear();
            foreach (var p in working.Properties) list.Items.Add(p);
            list.DisplayMember = nameof(PropertyDef.Display);
            if (list.Items.Count > 0 && list.SelectedIndex < 0) list.SelectedIndex = 0;
        }

        private void LoadSelected()
        {
            if (list.SelectedItem is PropertyDef d)
            {
                txtKey.Text = d.Key; txtTitle.Text = d.Title;
                txtEnum.Text = string.Join(',', d.EnumValues ?? new List<string>());
                txtDefault.Text = d.Default ?? "";
            }
        }

        private void SaveCurrent()
        {
            var key = txtKey.Text.Trim(); if (string.IsNullOrEmpty(key)) { MessageBox.Show("key 不能为空"); return; }
            var def = working.Properties.FirstOrDefault(p => p.Key.Equals(key, StringComparison.OrdinalIgnoreCase));
            if (def == null)
            {
                def = new PropertyDef { Key = key, Scope = "column", Kind = "enum" };
                working.Properties.Add(def);
            }
            def.Title = txtTitle.Text.Trim();
            def.Scope = "column";
            def.Kind  = string.IsNullOrEmpty(def.Kind) ? "enum" : def.Kind;
            def.EnumValues = txtEnum.Text.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                         .Select(s => s.Trim()).Distinct().ToList();
            def.Default = txtDefault.Text.Trim();
            RefreshList();
        }
    }

    public class AssignPropertyForm : Form
    {
        private readonly ListBox list;
        private readonly Panel editor;
        private readonly Button btnOk, btnCancel;
        private readonly List<PropertyDef> defs;
        private readonly Dictionary<string, PropertyKV> map;
        private PropertyDef? currentDef;
        private Control? currentInput;

        public List<PropertyKV> Result => map.Values.ToList();

        public AssignPropertyForm(List<PropertyDef> definitions, List<PropertyKV> existing)
        {
            Text = "设置列属性"; Width = 680; Height = 480;
            defs = definitions;
            map = existing.ToDictionary(p => p.PropertyName!, StringComparer.OrdinalIgnoreCase);

            list = new ListBox { Left = 10, Top = 10, Width = 240, Height = 380 };
            list.SelectedIndexChanged += (_, __) => { CommitEditor(); RenderEditor(); };
            Controls.Add(list);

            editor = new Panel { Left = 260, Top = 10, Width = 390, Height = 380, BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle };
            Controls.Add(editor);

            btnOk = new Button { Text = "确定", Left = Width - 200, Top = Height - 90, Width = 80, Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            btnCancel = new Button { Text = "取消", Left = Width - 110, Top = Height - 90, Width = 80, Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            btnOk.Click += (_, __) => { CommitEditor(); DialogResult = DialogResult.OK; Close(); };
            btnCancel.Click += (_, __) => { DialogResult = DialogResult.Cancel; Close(); };
            Controls.Add(btnOk); Controls.Add(btnCancel);

            foreach (var d in defs) list.Items.Add(d);
            list.DisplayMember = nameof(PropertyDef.Display);
            if (list.Items.Count > 0) list.SelectedIndex = 0;
        }

        private void RenderEditor()
        {
            editor.Controls.Clear();
            currentInput = null;
            currentDef = list.SelectedItem as PropertyDef;
            if (currentDef == null) return;

            var d = currentDef;
            editor.Controls.Add(new Label
            {
                Left = 10, Top = 10,
                Text = $"{d.Key}（{d.Title}）",
                AutoSize = true,
                Font = new System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Bold)
            });

            map.TryGetValue(d.Key, out var kv);
            var v = kv?.PropertyValue ?? d.Default ?? string.Empty;

            switch (d.Kind)
            {
                case "bool":
                    var chk = new CheckBox { Left = 10, Top = 40, Text = "true/false", Checked = v.Equals("true", StringComparison.OrdinalIgnoreCase) };
                    currentInput = chk; break;
                case "int":
                    var num = new NumericUpDown { Left = 10, Top = 40, Width = 120, Minimum = int.MinValue, Maximum = int.MaxValue, Value = int.TryParse(v, out var iv) ? iv : 0 };
                    currentInput = num; break;
                case "string":
                    var tb = new TextBox { Left = 10, Top = 40, Width = 320, Text = v };
                    currentInput = tb; break;
                default: // enum（默认）
                    var cb = new ComboBox { Left = 10, Top = 40, Width = 240, DropDownStyle = ComboBoxStyle.DropDownList };
                    cb.Items.AddRange((d.EnumValues ?? new List<string>()).ToArray());
                    if (!string.IsNullOrEmpty(v) && !cb.Items.Contains(v)) cb.Items.Add(v);
                    cb.SelectedItem = string.IsNullOrEmpty(v) ? (cb.Items.Count > 0 ? cb.Items[0] : null) : v;
                    currentInput = cb; break;
            }
            editor.Controls.Add(currentInput);
        }

        private void CommitEditor()
        {
            if (currentDef == null || currentInput == null) return;
            var d = currentDef;
            string val = d.Kind switch
            {
                "enum" => (currentInput as ComboBox)?.SelectedItem?.ToString() ?? string.Empty,
                "bool" => (currentInput as CheckBox)?.Checked == true ? "true" : "false",
                "int"  => ((int)(currentInput as NumericUpDown)!.Value).ToString(),
                _      => (currentInput as TextBox)?.Text ?? string.Empty
            };
            map[d.Key] = new PropertyKV(d.Key, val);
        }
    }

    public class ValidationResultForm : Form
    {
        private readonly DataGridView grid;
        private readonly Button btnExport, btnClose;

        public ValidationResultForm(string tableName, List<ValidationError> errors)
        {
            Text = errors.Count == 0 ? $"✅ {tableName}：未发现类型问题" : $"❌ {tableName}：共 {errors.Count} 处类型不匹配";
            Width = 900; Height = 520;

            grid = new DataGridView { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, AutoGenerateColumns = false };
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "行号", DataPropertyName = nameof(ValidationError.Row), Width = 70 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "列名", DataPropertyName = nameof(ValidationError.Column), Width = 160 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "期望类型", DataPropertyName = nameof(ValidationError.Expected), Width = 100 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "实际值", DataPropertyName = nameof(ValidationError.Value), AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "原因", DataPropertyName = nameof(ValidationError.Message), Width = 220 });

            grid.DataSource = errors;

            var bottom = new Panel { Dock = DockStyle.Bottom, Height = 44 };
            btnExport = new Button { Text = "导出CSV", Left = 10, Top = 8, Width = 90, Enabled = errors.Count > 0 };
            btnClose  = new Button { Text = "关闭", Left = 110, Top = 8, Width = 80 };

            btnExport.Click += (_, __) =>
            {
                using var sfd = new SaveFileDialog { Filter = "CSV 文件|*.csv", FileName = $"{tableName}.type_errors.csv" };
                if (sfd.ShowDialog(this) == DialogResult.OK)
                {
                    using var sw = new StreamWriter(sfd.FileName, false, new UTF8Encoding(false));
                    sw.WriteLine("Row,Column,Expected,Value,Message");
                    foreach (ValidationError e in (List<ValidationError>)grid.DataSource!)
                    {
                        string esc(string s) => "\"" + s.Replace("\"", "\"\"") + "\"";
                        sw.WriteLine($"{e.Row},{esc(e.Column)},{esc(e.Expected)},{esc(e.Value)},{esc(e.Message)}");
                    }
                    MessageBox.Show("已导出。");
                }
            };
            btnClose.Click += (_, __) => Close();

            bottom.Controls.Add(btnExport);
            bottom.Controls.Add(btnClose);
            Controls.Add(grid);
            Controls.Add(bottom);
            if (errors.Count == 0) btnExport.Enabled = false;
        }
    }
    public class ValidationError
    {
        public int Row { get; set; }
        public string Column { get; set; } = "";
        public string Expected { get; set; } = "";
        public string Value { get; set; } = "";
        public string Message { get; set; } = "";
    }

    // —— UI 状态 —— 
    public class UiState
    {
        [JsonProperty("splitter")] public int SplitterDistance { get; set; }
        [JsonProperty("col_widths")] public Dictionary<string, int> ColWidths { get; set; } = new();
    }
}
