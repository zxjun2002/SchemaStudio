using System.ComponentModel;
using System.Data;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using CsvHelper;
using System.Globalization;
using ExcelDataReader;

namespace SchemaStudio
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
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
        private readonly ToolTip tips = new ToolTip();

        private string? currentDir;
        private string? currentFilePath;
        private TableSchema? currentSchema;
        private PropertyRegistry registry = new();

        private static readonly string[] FileExts = new[] { ".csv", ".xlsx" };
        private static readonly string[] AllowedTypes = new[]
        {
            "int","long","float","double","bool","string","int[]","List<int>"
        };

        private string RegistryPath => currentDir == null ? "" : Path.Combine(currentDir!, "schema.registry.json");
        private string UiStatePath   => currentDir == null ? "" : Path.Combine(currentDir!, "schema.ui.json");

        public MainForm()
        {
            Text = "SchemaStudio — JSON 可视化编辑器";
            Width = 1280; Height = 760;

            split = new SplitContainer { Dock = DockStyle.Fill, SplitterDistance = 380 };
            topBar = new Panel { Dock = DockStyle.Top, Height = 40 };
            btnOpenDir = new Button { Text = "打开目录", Left = 10, Width = 90, Top = 8 };
            btnSave    = new Button { Text = "保存JSON", Left = 110, Width = 90, Top = 8 };
            btnCheck   = new Button { Text = "检查",     Left = 210, Width = 70, Top = 8 };
            btnNewTable= new Button { Text = "新建表",   Left = 290, Width = 80, Top = 8 };
            btnPropLib = new Button { Text = "自定义属性", Left = 380, Width = 110, Top = 8 };
            lblDir = new Label { Text = "未选择目录", AutoSize = true, Left = 510, Top = 12 };

            btnOpenDir.Click += (_, __) => ChooseDir();
            btnSave.Click    += (_, __) => SaveSchema();
            btnCheck.Click   += (_, __) => ValidateCurrent();
            btnNewTable.Click+= (_, __) => CreateNewTable();
            btnPropLib.Click += (_, __) => OpenPropertyLibrary();

            tips.SetToolTip(btnCheck, "校验当前表数据与配置的数值类型是否匹配");
            tips.SetToolTip(btnNewTable, "在选定目录下创建CSV并生成同名 .schema.json");
            tips.SetToolTip(btnPropLib, "编辑列自定义属性库(根目录 schema.registry.json，团队共享)");

            topBar.Controls.AddRange(new Control[] { btnOpenDir, btnSave, btnCheck, btnNewTable, btnPropLib, lblDir });
            Controls.Add(split);
            Controls.Add(topBar);

            // 左：目录树（递归）
            tree = new TreeView { Dock = DockStyle.Fill, HideSelection = false, ShowLines = true };
            tree.AfterSelect += (_, e) => TryLoadNodeFile(e.Node);
            tree.NodeMouseDoubleClick += (_, e) => TryLoadNodeFile(e.Node);
            split.Panel1.Controls.Add(tree);

            // 右：Grid
            grid = new DataGridView { Dock = DockStyle.Fill, AutoGenerateColumns = false, AllowUserToAddRows = false, RowHeadersVisible = false };
            BuildGridColumns();
            grid.DataSource = rows;
            grid.CellContentClick += Grid_CellContentClick;
            grid.ColumnWidthChanged += (_, __) => SaveUiState();
            split.SplitterMoved += (_, __) => SaveUiState();
            split.Panel2.Controls.Add(grid);

            this.FormClosing += (_, __) => SaveUiState();
        }

        private void BuildGridColumns()
        {
            grid.Columns.Clear();

            var colName = new DataGridViewTextBoxColumn
            {
                Name = "col_name",
                HeaderText = "列名",
                DataPropertyName = nameof(ColumnRow.ColName),
                ReadOnly = true,
                Width = 260
            };
            grid.Columns.Add(colName);

            var typeCol = new DataGridViewComboBoxColumn
            {
                Name = "value_type",
                HeaderText = "数值类型",
                DataPropertyName = nameof(ColumnRow.ValueType),
                Width = 140,
                DataSource = AllowedTypes
            };
            grid.Columns.Add(typeCol);

            var pkCol = new DataGridViewTextBoxColumn
            {
                Name = "pk_order",
                HeaderText = "主键序列",
                DataPropertyName = nameof(ColumnRow.PkOrderText),
                Width = 150
            };
            grid.Columns.Add(pkCol);

            var btn = new DataGridViewButtonColumn
            {
                Name = "prop_btn",
                HeaderText = "列属性",
                Text = "属性…",
                UseColumnTextForButtonValue = true,
                Width = 80
            };
            grid.Columns.Add(btn);

            var summary = new DataGridViewTextBoxColumn
            {
                Name = "prop_summary",
                HeaderText = "属性摘要",
                DataPropertyName = nameof(ColumnRow.PropsSummary),
                ReadOnly = true,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            };
            grid.Columns.Add(summary);
        }

        private void ApplyUiState()
        {
            if (string.IsNullOrEmpty(UiStatePath) || !File.Exists(UiStatePath)) return;
            try
            {
                var state = JsonSerializer.Deserialize<UiState>(File.ReadAllText(UiStatePath));
                if (state == null) return;
                if (state.SplitterDistance > 0) split.SplitterDistance = state.SplitterDistance;
                foreach (DataGridViewColumn c in grid.Columns)
                {
                    if (state.ColWidths.TryGetValue(c.Name, out var w) && w > 0)
                        c.Width = w;
                }
            }
            catch { /* ignore */ }
        }

        private void SaveUiState()
        {
            if (string.IsNullOrEmpty(UiStatePath)) return;
            try
            {
                var st = new UiState { SplitterDistance = split.SplitterDistance, ColWidths = new Dictionary<string, int>() };
                foreach (DataGridViewColumn c in grid.Columns)
                {
                    st.ColWidths[c.Name] = c.Width;
                }
                File.WriteAllText(UiStatePath, JsonSerializer.Serialize(st, new JsonSerializerOptions { WriteIndented = true }));
            }
            catch { /* ignore */ }
        }

        private void Grid_CellContentClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            var btnColIndex = grid.Columns["prop_btn"]!.Index;
            if (e.ColumnIndex == btnColIndex && grid.Rows[e.RowIndex].DataBoundItem is ColumnRow row)
            {
                if (currentSchema == null) return;
                var col = currentSchema.Columns!.First(c => c.ColName == row.ColName);
                var defs = registry.Properties.Where(d => d.Scope == "column" || d.Scope == "both").ToList();

                using var dlg = new AssignPropertyForm(defs, col.ColProperties ?? new List<PropertyKV>());
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    col.ColProperties = dlg.Result;
                    row.PropsSummary = Summarize(col.ColProperties);
                    grid.InvalidateRow(e.RowIndex);
                }
            }
        }

        private string Summarize(List<PropertyKV>? list)
        {
            if (list == null || list.Count == 0) return "";
            return string.Join("; ", list.Select(p => $"{p.PropertyName}={p.PropertyValue}"));
        }

        private void ChooseDir()
        {
            using var fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                currentDir = fbd.SelectedPath;
                lblDir.Text = currentDir;
                LoadRegistry();
                BuildDirectoryTree();
                ApplyUiState(); // 应用用户上次的列宽/分割条比例
            }
        }

        private void LoadRegistry()
        {
            try
            {
                registry = File.Exists(RegistryPath)
                    ? JsonSerializer.Deserialize<PropertyRegistry>(File.ReadAllText(RegistryPath), JsonOpt()) ?? new PropertyRegistry()
                    : new PropertyRegistry();

                EnsureProp("cs_type", "终端类型", scope: "column", kind: "enum",
                    values: new[] { "ALL", "C", "S" }, defVal: "ALL");
                EnsureProp("translate_mode", "翻译类型", scope: "column", kind: "enum",
                    values: new[] { "none", "auto", "manual" }, defVal: "none");

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
                registry.Properties.Add(new PropertyDef
                {
                    Key = key, Title = title, Scope = scope, Kind = kind,
                    EnumValues = values.ToList(), Default = defVal
                });
            }
        }

        private void SaveRegistry()
        {
            try
            {
                if (string.IsNullOrEmpty(RegistryPath)) return;
                var json = JsonSerializer.Serialize(registry, JsonOpt(true));
                File.WriteAllText(RegistryPath, json, new UTF8Encoding(false));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存属性库失败：{ex.Message}");
            }
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

        private void BuildDirectoryTree()
        {
            tree.BeginUpdate();
            try
            {
                tree.Nodes.Clear();
                if (currentDir == null || !Directory.Exists(currentDir)) return;

                var root = new TreeNode(Path.GetFileName(currentDir.TrimEnd(Path.DirectorySeparatorChar))) { Tag = currentDir };
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
                var node = new TreeNode(Path.GetFileName(sub)) { Tag = sub };
                parent.Nodes.Add(node);
                AddDirNodes(node, sub);
            }
            foreach (var file in Directory.GetFiles(dir)
                         .Where(p => FileExts.Contains(Path.GetExtension(p).ToLowerInvariant()))
                         .Where(p => !Path.GetFileName(p).StartsWith("~$"))
                         .OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
            {
                var node = new TreeNode(Path.GetFileName(file)) { Tag = file };
                parent.Nodes.Add(node);
            }
        }

        private void TryLoadNodeFile(TreeNode? node)
        {
            if (node?.Tag is string path)
            {
                var ext = Path.GetExtension(path).ToLowerInvariant();
                if (ext == ".csv" || ext == ".xlsx") LoadFile(path);
            }
        }

        private void LoadFile(string path)
        {
            currentFilePath = path;
            var tableName = Path.GetFileName(path);
            var jsonPath = GetSchemaJsonPath(path);

            var colNames = ReadHeaderColumns(path);
            if (colNames.Count == 0) { MessageBox.Show("未能读取到列名。请确认第1行是字段名。"); return; }

            TableSchema schema = File.Exists(jsonPath)
                ? (JsonSerializer.Deserialize<TableSchema>(File.ReadAllText(jsonPath), JsonOpt()) ?? new TableSchema { TableName = tableName })
                : new TableSchema { TableName = tableName };
            schema.TableName ??= tableName;

            var map = schema.Columns?.ToDictionary(c => c.ColName ?? "", StringComparer.OrdinalIgnoreCase) ?? new();
            var newCols = new List<ColumnSchema>();
            for (int i = 0; i < colNames.Count; i++)
            {
                var name = colNames[i];
                if (map.TryGetValue(name, out var c))
                {
                    c.ColSort = i;
                    newCols.Add(c);
                }
                else
                {
                    newCols.Add(new ColumnSchema
                    {
                        ColName = name,
                        ValueType = "string",
                        PkOrder = -1,
                        PrimaryKey = false,
                        ColSort = i,
                        ColProperties = new List<PropertyKV>()
                    });
                }
            }
            schema.Columns = newCols;
            currentSchema = schema;

            rows.Clear();
            foreach (var c in schema.Columns)
            {
                rows.Add(ColumnRow.FromSchema(c));
            }
            Text = $"SchemaStudio — {tableName}";
        }

        private void SaveSchema()
        {
            if (currentSchema == null || currentFilePath == null) return;

            foreach (var r in rows)
            {
                var c = currentSchema.Columns!.First(x => x.ColName == r.ColName);
                c.ValueType = r.ValueType;
                c.PkOrder = r.PkOrder; // 已由 PkOrderText 同步
                c.PrimaryKey = r.PkOrder >= 0;
            }

            var jsonPath = GetSchemaJsonPath(currentFilePath);
            var json = JsonSerializer.Serialize(currentSchema, JsonOpt(true));
            File.WriteAllText(jsonPath, json, new UTF8Encoding(false));
            MessageBox.Show($"已保存：{jsonPath}", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // —— 校验 —— //
        private void ValidateCurrent()
        {
            if (currentFilePath == null || currentSchema == null)
            {
                MessageBox.Show("请先在左侧选择一张表。");
                return;
            }

            var headers = ReadHeaderColumns(currentFilePath);
            if (headers.Count == 0) { MessageBox.Show("未能读取到列名。"); return; }
            var data = ReadAllRows(currentFilePath, headers.Count);

            var headerIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < headers.Count; i++) headerIndex[headers[i]] = i;

            var errors = new List<ValidationError>();
            for (int rowIdx = 0; rowIdx < data.Count; rowIdx++)
            {
                var row = data[rowIdx];
                foreach (var col in currentSchema.Columns!)
                {
                    var name = col.ColName ?? "";
                    if (!headerIndex.TryGetValue(name, out var colIdx)) continue;
                    var val = colIdx < row.Length ? row[colIdx] ?? "" : "";
                    if (!IsValid(val, col.ValueType ?? "string", out var reason))
                    {
                        errors.Add(new ValidationError
                        {
                            Row = rowIdx + 2,
                            Column = name,
                            Value = val,
                            Expected = col.ValueType ?? "string",
                            Message = reason
                        });
                    }
                }
            }

            using var dlg = new ValidationResultForm(Path.GetFileName(currentFilePath), errors);
            dlg.ShowDialog(this);
        }

        private static bool IsValid(string value, string type, out string reason)
        {
            reason = "";
            if (string.IsNullOrWhiteSpace(value)) return true;

            var t = (type ?? "string").Trim();
            NumberStyles ns = NumberStyles.Integer;
            var ic = CultureInfo.InvariantCulture;

            bool ParseBool(string s)
            {
                s = s.Trim().ToLowerInvariant();
                if (bool.TryParse(s, out _)) return true;
                return s is "1" or "0" or "y" or "n" or "yes" or "no" or "是" or "否" or "真" or "假";
            }

            bool ParseIntSeq(string s)
            {
                var parts = s.Split(new[] { ',', '，', ';', '|', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var p in parts)
                    if (!int.TryParse(p.Trim(), ns, ic, out _)) return false;
                return true;
            }

            switch (t)
            {
                case "int":    if (!int.TryParse(value, ns, ic, out _)) { reason = "不是有效的 int"; return false; } return true;
                case "long":   if (!long.TryParse(value, ns, ic, out _)) { reason = "不是有效的 long"; return false; } return true;
                case "float":  if (!float.TryParse(value, NumberStyles.Float, ic, out _)) { reason = "不是有效的 float"; return false; } return true;
                case "double": if (!double.TryParse(value, NumberStyles.Float, ic, out _)) { reason = "不是有效的 double"; return false; } return true;
                case "bool":   if (!ParseBool(value)) { reason = "不是有效的 bool(true/false/1/0/是/否)"; return false; } return true;
                case "string": return true;
                case "int[]":
                case "List<int>":
                    if (!ParseIntSeq(value)) { reason = "不是有效的整型数组(分隔符支持 , ; | 空格)"; return false; }
                    return true;
                default:       return true;
            }
        }

        private static List<string[]> ReadAllRows(string path, int expectedCols)
        {
            var ext = Path.GetExtension(path).ToLowerInvariant();
            if (ext == ".csv") return ReadCsvRows(path, expectedCols);
            if (ext == ".xlsx") return ReadXlsxRows(path, expectedCols);
            return new();
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
                for (int i = 0; i < expectedCols; i++)
                {
                    csv.TryGetField<string>(i, out var v);
                    row[i] = v?.Trim() ?? "";
                }
                list.Add(row);
            }
            return list;
        }

        private static List<string[]> ReadXlsxRows(string path, int expectedCols)
        {
            var list = new List<string[]>();
            using var fs = File.OpenRead(path);
            using var reader = ExcelDataReader.ExcelReaderFactory.CreateReader(fs);
            var ds = reader.AsDataSet();
            if (ds.Tables.Count == 0) return list;
            var table = ds.Tables[0];
            if (table.Rows.Count <= 1) return list;

            for (int r = 1; r < table.Rows.Count; r++)
            {
                var dr = table.Rows[r];
                var row = new string[expectedCols];
                for (int c = 0; c < expectedCols; c++)
                {
                    row[c] = dr[c]?.ToString()?.Trim() ?? "";
                }
                list.Add(row);
            }
            return list;
        }

        private static string GetSchemaJsonPath(string tablePath)
            => Path.Combine(Path.GetDirectoryName(tablePath)!, Path.GetFileName(tablePath) + ".schema.json");

        private static List<string> ReadHeaderColumns(string path)
        {
            var ext = Path.GetExtension(path).ToLowerInvariant();
            if (ext == ".csv") return ReadCsvHeader(path);
            if (ext == ".xlsx") return ReadXlsxHeader(path);
            return new();
        }

        private static List<string> ReadCsvHeader(string path)
        {
            using var reader = new StreamReader(path, DetectEncoding(path));
            using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
            if (!csv.Read()) return new();
            var header = new List<string>();
            for (int i = 0; csv.TryGetField<string>(i, out var val); i++)
                header.Add((val ?? $"col_{i}").Trim());
            return header;
        }

        private static List<string> ReadXlsxHeader(string path)
        {
            using var fs = File.OpenRead(path);
            using var reader = ExcelReaderFactory.CreateReader(fs);
            var ds = reader.AsDataSet();
            if (ds.Tables.Count == 0) return new();
            var table = ds.Tables[0];
            if (table.Rows.Count == 0) return new();
            var row = table.Rows[0];
            var list = new List<string>();
            for (int i = 0; i < table.Columns.Count; i++)
            {
                var val = row[i]?.ToString() ?? $"col_{i}";
                list.Add(val.Trim());
            }
            return list;
        }

        private static Encoding DetectEncoding(string filePath)
        {
            var bytes = File.ReadAllBytes(filePath);
            try { Encoding.UTF8.GetString(bytes); return new UTF8Encoding(false); }
            catch { return Encoding.UTF8; }
        }

        private static JsonSerializerOptions JsonOpt(bool indented = false) => new JsonSerializerOptions
        {
            WriteIndented = indented,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
        };

        // —— 新建表 —— //
        private void CreateNewTable()
        {
            if (currentDir == null) { MessageBox.Show("请先打开根目录"); return; }
            var dirs = GetAllDirs(currentDir).ToList();
            using var dlg = new NewTableForm(currentDir, dirs, registry, AllowedTypes);
            if (dlg.ShowDialog(this) == DialogResult.OK)
            {
                var destDir = dlg.ResultDir!;
                var fileName = dlg.ResultFileName!; // 不带扩展
                var path = Path.Combine(destDir, fileName + ".csv");
                if (File.Exists(path))
                {
                    MessageBox.Show("同名CSV已存在，已取消。");
                    return;
                }

                // 写 CSV（UTF-8 BOM）
                Directory.CreateDirectory(destDir);
                using (var sw = new StreamWriter(path, false, new UTF8Encoding(true)))
                {
                    sw.WriteLine(string.Join(",", dlg.Columns.Select(c => CsvEscape(c.ColName))));
                }

                // 生成 schema.json
                var schema = new TableSchema
                {
                    TableName = Path.GetFileName(path),
                    Columns = dlg.Columns.Select((c, idx) => new ColumnSchema
                    {
                        ColName = c.ColName,
                        ValueType = c.ValueType,
                        PkOrder = c.PkOrder,
                        PrimaryKey = c.PkOrder >= 0,
                        ColSort = idx,
                        ColProperties = c.ColProperties?.ToList() ?? new List<PropertyKV>()
                    }).ToList()
                };
                var json = JsonSerializer.Serialize(schema, JsonOpt(true));
                File.WriteAllText(GetSchemaJsonPath(path), json, new UTF8Encoding(false));

                // 刷新树并选中
                BuildDirectoryTree();
                SelectPathInTree(path);
                MessageBox.Show($"已创建：\n{path}\n{path}.schema.json");
            }

            static string CsvEscape(string s)
            {
                if (s.Contains('"') || s.Contains(',') || s.Contains('\n') || s.Contains('\r'))
                    return "\"" + s.Replace("\"", "\"\"") + "\"";
                return s;
            }
        }

        private IEnumerable<string> GetAllDirs(string root)
        {
            yield return root;
            foreach (var d in Directory.GetDirectories(root))
            {
                foreach (var x in GetAllDirs(d)) yield return x;
            }
        }

        private void SelectPathInTree(string fullPath)
        {
            foreach (TreeNode root in tree.Nodes)
            {
                var node = FindNode(root, fullPath);
                if (node != null)
                {
                    tree.SelectedNode = node;
                    node.EnsureVisible();
                    TryLoadNodeFile(node);
                    break;
                }
            }

            TreeNode? FindNode(TreeNode n, string path)
            {
                if (n.Tag as string == path) return n;
                foreach (TreeNode c in n.Nodes)
                {
                    var got = FindNode(c, path);
                    if (got != null) return got;
                }
                return null;
            }
        }
    }

    // —— 新建表对话框 —— //
    public class NewTableForm : Form
    {
        private readonly ComboBox cmbDir;
        private readonly TextBox txtName;
        private readonly DataGridView grid;
        private readonly Button btnAdd, btnDel, btnOk, btnCancel;
        private readonly BindingList<NewColRow> rows = new();
        private readonly PropertyRegistry registry;
        private readonly string[] allowedTypes;

        public string? ResultDir { get; private set; }
        public string? ResultFileName { get; private set; } // 不带扩展
        public IEnumerable<NewColRow> Columns => rows;

        public NewTableForm(string root, List<string> allDirs, PropertyRegistry reg, string[] allowedTypes)
        {
            Text = "新建表";
            Width = 900; Height = 560;
            registry = reg;
            this.allowedTypes = allowedTypes;

            var lblDir = new Label { Left = 12, Top = 14, Text = "保存到目录：", AutoSize = true };
            cmbDir = new ComboBox { Left = 90, Top = 10, Width = 560, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbDir.Items.AddRange(allDirs.Select(d => d.Replace(root, "").TrimStart(Path.DirectorySeparatorChar)).Select(s => string.IsNullOrEmpty(s) ? "(根目录)" : s).ToArray());
            cmbDir.SelectedIndex = 0;

            var lblName = new Label { Left = 12, Top = 45, Text = "文件名(.csv)：", AutoSize = true };
            txtName = new TextBox { Left = 90, Top = 41, Width = 260, Text = "new_table" };

            grid = new DataGridView { Left = 12, Top = 80, Width = 860, Height = 380, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom, AutoGenerateColumns = false, AllowUserToAddRows = false };
            BuildColsGrid();

            btnAdd = new Button { Text = "添加列", Left = 12, Top = 470, Width = 80, Anchor = AnchorStyles.Bottom | AnchorStyles.Left };
            btnDel = new Button { Text = "删除选中", Left = 100, Top = 470, Width = 90, Anchor = AnchorStyles.Bottom | AnchorStyles.Left };
            btnOk  = new Button { Text = "确定", Left = 700, Top = 470, Width = 80, Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            btnCancel = new Button { Text = "取消", Left = 792, Top = 470, Width = 80, Anchor = AnchorStyles.Bottom | AnchorStyles.Right };

            btnAdd.Click += (_, __) => rows.Add(new NewColRow { ColName = $"col_{rows.Count+1}", ValueType = "string", PkOrder = -1 });
            btnDel.Click += (_, __) => { foreach (DataGridViewRow r in grid.SelectedRows) if (r.DataBoundItem is NewColRow x) rows.Remove(x); };
            btnOk.Click  += (_, __) => OnOk(root, allDirs);
            btnCancel.Click += (_, __) => { DialogResult = DialogResult.Cancel; Close(); };

            Controls.AddRange(new Control[] { lblDir, cmbDir, lblName, txtName, grid, btnAdd, btnDel, btnOk, btnCancel });

            // 默认两列示例
            rows.Add(new NewColRow { ColName = "id", ValueType = "int", PkOrder = 1 });
            rows.Add(new NewColRow { ColName = "desc", ValueType = "string", PkOrder = -1 });
        }

        private void BuildColsGrid()
        {
            grid.Columns.Clear();
            grid.DataSource = rows;

            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "列名", DataPropertyName = nameof(NewColRow.ColName), Width = 240 });
            grid.Columns.Add(new DataGridViewComboBoxColumn { HeaderText = "数值类型", DataPropertyName = nameof(NewColRow.ValueType), Width = 140, DataSource = allowedTypes });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "主键序列", DataPropertyName = nameof(NewColRow.PkOrderText), Width = 120 });

            var btn = new DataGridViewButtonColumn { HeaderText = "列属性", Text = "属性…", UseColumnTextForButtonValue = true, Width = 80 };
            grid.Columns.Add(btn);
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "属性摘要", DataPropertyName = nameof(NewColRow.PropsSummary), AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, ReadOnly = true });

            grid.CellContentClick += (s, e) =>
            {
                if (e.RowIndex < 0) return;
                if (grid.Columns[e.ColumnIndex] is DataGridViewButtonColumn && grid.Rows[e.RowIndex].DataBoundItem is NewColRow row)
                {
                    var defs = registry.Properties.Where(d => d.Scope == "column" || d.Scope == "both").ToList();
                    using var dlg = new AssignPropertyForm(defs, row.ColProperties ?? new List<PropertyKV>());
                    if (dlg.ShowDialog(this) == DialogResult.OK)
                    {
                        row.ColProperties = dlg.Result;
                        row.PropsSummary = string.Join("; ", row.ColProperties.Select(p => $"{p.PropertyName}={p.PropertyValue}"));
                        grid.InvalidateRow(e.RowIndex);
                    }
                }
            };
        }

        private void OnOk(string root, List<string> allDirs)
        {
            var name = (txtName.Text ?? "").Trim();
            if (string.IsNullOrEmpty(name)) { MessageBox.Show("文件名不能为空"); return; }
            if (name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) { MessageBox.Show("文件名包含非法字符"); return; }

            // 解析目录
            var rel = cmbDir.SelectedItem!.ToString()!;
            string destDir = rel == "(根目录)" ? root : Path.Combine(root, rel.TrimStart(Path.DirectorySeparatorChar));
            if (!allDirs.Contains(destDir)) { MessageBox.Show("目标目录无效"); return; }

            // 校验列
            if (rows.Count == 0) { MessageBox.Show("请至少添加一列"); return; }
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var c in rows)
            {
                if (string.IsNullOrWhiteSpace(c.ColName)) { MessageBox.Show("存在空的列名"); return; }
                if (!names.Add(c.ColName.Trim())) { MessageBox.Show($"列名重复：{c.ColName}"); return; }
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
            set
            {
                if (string.IsNullOrWhiteSpace(value)) PkOrder = -1;
                else if (int.TryParse(value.Trim(), out var v)) PkOrder = v;
                else PkOrder = -1;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(PkOrderText)));
            }
        }
    }

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
            set
            {
                if (string.IsNullOrWhiteSpace(value)) PkOrder = -1;
                else if (int.TryParse(value.Trim(), out var v)) PkOrder = v;
                else PkOrder = -1;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(PkOrderText)));
            }
        }

        public static ColumnRow FromSchema(ColumnSchema c) => new ColumnRow
        {
            ColName = c.ColName ?? "",
            ValueType = string.IsNullOrWhiteSpace(c.ValueType) ? "string" : c.ValueType!,
            PkOrder = c.PkOrder,
            PropsSummary = c.ColProperties == null || c.ColProperties.Count == 0
                ? ""
                : string.Join("; ", c.ColProperties.Select(p => $"{p.PropertyName}={p.PropertyValue}"))
        };
    }

    public class TableSchema
    {
        [JsonPropertyName("table_name")] public string? TableName { get; set; }
        [JsonPropertyName("columns")] public List<ColumnSchema>? Columns { get; set; }
    }

    public class ColumnSchema
    {
        [JsonPropertyName("col_name")] public string? ColName { get; set; }
        [JsonPropertyName("value_type")] public string? ValueType { get; set; }
        [JsonPropertyName("pk_order")] public int PkOrder { get; set; } = -1;
        [JsonPropertyName("primary_key")] public bool PrimaryKey { get; set; }
        [JsonPropertyName("col_sort")] public int ColSort { get; set; }
        [JsonPropertyName("col_properties")] public List<PropertyKV>? ColProperties { get; set; }
    }

    public class PropertyKV
    {
        public PropertyKV() { }
        public PropertyKV(string name, string value) { PropertyName = name; PropertyValue = value; }
        [JsonPropertyName("property_name")] public string? PropertyName { get; set; }
        [JsonPropertyName("property_value")] public string? PropertyValue { get; set; }
    }

    public class PropertyRegistry
    {
        [JsonPropertyName("version")] public int Version { get; set; } = 1;
        [JsonPropertyName("properties")] public List<PropertyDef> Properties { get; set; } = new();
    }

    public class PropertyDef
    {
        [JsonPropertyName("key")] public string Key { get; set; } = string.Empty;
        [JsonPropertyName("title")] public string Title { get; set; } = string.Empty;
        [JsonPropertyName("scope")] public string Scope { get; set; } = "column";
        [JsonPropertyName("kind")] public string Kind { get; set; } = "enum"; // 默认枚举（仍兼容其它）
        [JsonPropertyName("values")] public List<string>? EnumValues { get; set; } = new();
        [JsonPropertyName("default")] public string? Default { get; set; }

        [JsonIgnore] public string Display => string.IsNullOrEmpty(Title) ? Key : $"{Key}（{Title}）";
        public override string ToString() => Display;
    }

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
            working = JsonSerializer.Deserialize<PropertyRegistry>(JsonSerializer.Serialize(reg))!; // 深拷贝

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

            editor = new Panel { Left = 260, Top = 10, Width = 390, Height = 380, BorderStyle = BorderStyle.FixedSingle };
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
                _       => (currentInput as TextBox)?.Text ?? string.Empty
            };
            map[d.Key] = new PropertyKV(d.Key, val);
        }
    }

    // —— 校验结果窗口 —— //
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

    // —— UI 状态（列宽/分割条） —— //
    public class UiState
    {
        [JsonPropertyName("splitter")] public int SplitterDistance { get; set; }
        [JsonPropertyName("col_widths")] public Dictionary<string, int> ColWidths { get; set; } = new();
    }
}
