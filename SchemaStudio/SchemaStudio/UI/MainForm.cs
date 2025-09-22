using System.ComponentModel;
using System.Text;
using System.Text.Json;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SchemaStudio.Models;
using SchemaStudio.Services;

namespace SchemaStudio.UI
{
    public class MainForm : Form
    {
        private readonly SplitContainer split;
        private readonly Panel topBar;
        private readonly Button btnOpenDir, btnSave, btnCheck, btnNewTable, btnPropLib;
        private readonly Button btnTableProps;
        private readonly Label lblTablePropSummary;
        private readonly Label lblDir;
        private readonly TreeView tree;
        private readonly DataGridView grid;
        private readonly BindingList<ColumnRow> rows = new();
        private readonly ToolTip tips = new();

        private string? currentDir;
        private NodeTag? currentTag;
        private TableSchema? currentSchema;
        private WorkbookSchema? currentWorkbook;
        private PropertyRegistry registry = new();

        private string RegistryPath => currentDir == null ? "" : Path.Combine(currentDir!, "schema.registry.json");

        public MainForm()
        {
            Text = "SchemaStudio — JSON 可视化编辑器";
            Width = 1280; Height = 760;

            //DPI 自适应 + 用系统消息字体
            AutoScaleMode = AutoScaleMode.Dpi;
            Font = SystemFonts.MessageBoxFont;

            // 工具：创建自适应按钮
            Button MakeBtn(string text) => new Button {
                Text = text,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = new Padding(0, 0, 8, 0)   // 按钮间距
            };

            // 顶部栏用 FlowLayoutPanel
            topBar = new FlowLayoutPanel {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = false,
                Padding = new Padding(8, 6, 8, 6)
            };

            btnOpenDir  = MakeBtn("打开目录");
            btnSave     = MakeBtn("保存JSON");
            btnCheck    = MakeBtn("检查");
            btnNewTable = MakeBtn("新建表");
            btnPropLib  = MakeBtn("自定义属性");
            btnTableProps = MakeBtn("表属性…");
            lblTablePropSummary = new Label { Text = "表属性：—", AutoSize = true, Margin = new Padding(12, 6, 0, 0) };

            lblDir = new Label { Text = "未选择目录", AutoSize = true, Margin = new Padding(12, 6, 0, 0) };
            
            btnOpenDir.Click += (_, __) => ChooseDir();
            btnSave.Click    += (_, __) => SaveSchema();
            btnCheck.Click   += (_, __) => ValidateCurrent();
            btnNewTable.Click+= (_, __) => CreateNewTable();
            btnPropLib.Click += (_, __) => OpenPropertyLibrary();
            btnTableProps.Click += (_, __) => EditCurrentTableProps();

            tips.SetToolTip(btnCheck,   "校验当前表数据与配置的数值类型是否匹配");
            tips.SetToolTip(btnNewTable,"新建 CSV / XLSX / XLS（Excel 可创建多个工作表）");
            tips.SetToolTip(btnPropLib, "编辑列自定义属性库(根目录 schema.registry.json)");
            tips.SetToolTip(btnTableProps, "为当前表（CSV）或当前工作表（Excel Sheet）设置表属性");

            topBar.Controls.AddRange(new Control[] { btnOpenDir, btnSave, btnCheck, btnNewTable, btnPropLib, btnTableProps, lblDir, lblTablePropSummary});
            split = new SplitContainer { Dock = DockStyle.Fill, SplitterDistance = UiDefaults.LeftTreeWidth };

            tree = new TreeView { Dock = DockStyle.Fill, HideSelection = false, ShowLines = true };
            tree.BeforeExpand += (_, e) => { if (e.Node?.Tag is NodeTag t && t.IsExcelFile && e.Node.Nodes.Count == 1 && e.Node.Nodes[0].Tag == null) AddSheetChildren(e.Node, t.Path); };
            tree.AfterSelect += (_, e) => TryLoadNode(e.Node);
            tree.NodeMouseDoubleClick += (_, e) => TryLoadNode(e.Node);
            split.Panel1.Controls.Add(tree);

            grid = new DataGridView {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                AllowUserToAddRows = false,
                RowHeadersVisible = false,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
            };
            BuildGridColumns();
            grid.DataSource = rows;
            grid.CellContentClick += Grid_CellContentClick;
            split.Panel2.Controls.Add(grid);

            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2 };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            root.Controls.Add(topBar, 0, 0);
            root.Controls.Add(split, 0, 1);
            Controls.Add(root);
        }

        private void BuildGridColumns()
        {
            grid.Columns.Clear();
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name="col_name", HeaderText="列名", DataPropertyName=nameof(ColumnRow.ColName), ReadOnly=true, Width=UiDefaults.GridColWidths["col_name"] });
            grid.Columns.Add(new DataGridViewComboBoxColumn { Name="value_type", HeaderText="数值类型", DataPropertyName=nameof(ColumnRow.ValueType), Width=UiDefaults.GridColWidths["value_type"], DataSource=UiDefaults.AllowedTypes });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name="pk_order", HeaderText="主键序列", DataPropertyName=nameof(ColumnRow.PkOrderText), Width=UiDefaults.GridColWidths["pk_order"] });
            grid.Columns.Add(new DataGridViewButtonColumn { Name="prop_btn", HeaderText="列属性", Text="属性…", UseColumnTextForButtonValue=true, Width=UiDefaults.GridColWidths["prop_btn"] });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name="prop_summary", HeaderText="属性摘要", DataPropertyName=nameof(ColumnRow.PropsSummary), ReadOnly=true, Width=UiDefaults.GridColWidths["prop_summary"] });
        }

        // ======= 打开目录/属性库 ======= //
        private void ChooseDir()
        {
            using var fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() != DialogResult.OK) return;

            currentDir = fbd.SelectedPath;
            lblDir.Text = currentDir;
            LoadRegistry();
            BuildDirectoryTree();
        }
        private void LoadRegistry()
        {
            try
            {
                registry = DataIO.LoadRegistry(RegistryPath);
                // 默认分表设置
                EnsureProp("cs_type", "终端类型", "column", "enum", new[] { "ALL","C","S" }, "ALL");
                DataIO.SaveRegistry(RegistryPath, registry);
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
                registry.Properties.Add(new PropertyDef { Key = key, Title = title, Scope = scope, Kind = kind, EnumValues = values.ToList(), Default = defVal });
        }
        private void OpenPropertyLibrary()
        {
            if (currentDir == null) { MessageBox.Show("请先打开目录"); return; }
            using var dlg = new PropertyLibraryForm(registry);
            if (dlg.ShowDialog(this) == DialogResult.OK)
            {
                registry = dlg.Result;
                DataIO.SaveRegistry(RegistryPath, registry);
            }
        }
        private void EditCurrentTableProps()
        {
            if (currentTag == null)
            {
                MessageBox.Show("请先选择 CSV 或 Excel 的某个工作表。");
                return;
            }
            if (currentSchema == null)
            {
                MessageBox.Show("未加载到当前表结构。");
                return;
            }

            // 属性库共用列属性定义即可（scope=column 或 both）
            var defs = registry.Properties.Where(d => d.Scope == "column" || d.Scope == "both").ToList();

            var existing = currentSchema.TableProperties ?? new List<PropertyKV>();
            using var dlg = new AssignPropertyForm(defs, existing);
            if (dlg.ShowDialog(this) == DialogResult.OK)
            {
                currentSchema.TableProperties = dlg.Result ?? new List<PropertyKV>();
                RefreshTablePropSummary();
            }
        }

        // ======= 目录树（懒加载 Sheet） ======= //
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
                         .Where(p => UiDefaults.FileExts.Contains(Path.GetExtension(p).ToLowerInvariant()))
                         .Where(p => !Path.GetFileName(p).StartsWith("~$"))
                         .OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
            {
                var ext = Path.GetExtension(file).ToLowerInvariant();
                var node = new TreeNode(Path.GetFileName(file)) { Tag = new NodeTag(file, null) };
                parent.Nodes.Add(node);
                if (ext is ".xlsx" or ".xls")
                    node.Nodes.Add(new TreeNode("(展开查看工作表)"));
            }
        }
        private void AddSheetChildren(TreeNode fileNode, string excelPath)
        {
            fileNode.Nodes.Clear();
            try
            {
                using var fs = File.OpenRead(excelPath);
                IWorkbook wb = Path.GetExtension(excelPath).Equals(".xls", StringComparison.OrdinalIgnoreCase) ? new HSSFWorkbook(fs) : new XSSFWorkbook(fs);
                for (int i = 0; i < wb.NumberOfSheets; i++)
                    fileNode.Nodes.Add(new TreeNode(wb.GetSheetName(i)) { Tag = new NodeTag(excelPath, wb.GetSheetName(i)) });
            }
            catch (Exception ex) { fileNode.Nodes.Add(new TreeNode($"(读取失败: {ex.Message})")); }
        }
        private void TryLoadNode(TreeNode? node)
        {
            if (node?.Tag is not NodeTag tag) return;
            currentTag = tag; currentSchema = null; currentWorkbook = null;

            if (tag.IsExcelFile) { rows.Clear(); Text = $"SchemaStudio — {Path.GetFileName(tag.Path)}"; return; }

            if (tag.IsExcelSheet) LoadExcelSheet(tag.Path, tag.Sheet!);
            else if (tag.IsCsv)   LoadCsv(tag.Path);
        }

        // ======= 加载 CSV/Excel Sheet ======= //
        private void LoadCsv(string path)
        {
            var tableName = Path.GetFileName(path);
            var jsonPath  = DataIO.GetSchemaJsonPath(path);

            var colNames = DataIO.ReadCsvHeader(path);
            if (colNames.Count == 0) { MessageBox.Show("未能读取到列名。请确认第1行是字段名。"); return; }

            var schema = File.Exists(jsonPath)
                ? JsonSerializer.Deserialize<TableSchema>(File.ReadAllText(jsonPath), DataIO.JsonOptions) ?? new TableSchema { TableName = tableName }
                : new TableSchema { TableName = tableName };

            var cols = MapColumns(colNames, schema.Columns);
            var dataRowsForGuess = DataIO.ReadCsvRows(path, colNames.Count);
            TypeInference.ApplyRecommendedTypes(cols, dataRowsForGuess);

            currentSchema = new TableSchema
            {
                TableName = tableName,
                Columns = cols,
                TableProperties = schema.TableProperties ?? new List<PropertyKV>()
            };
            rows.Clear(); foreach (var c in cols) rows.Add(ColumnRow.FromSchema(c));
            Text = $"SchemaStudio — {tableName}";
            RefreshTablePropSummary();
        }
        private void LoadExcelSheet(string path, string sheetName)
        {
            var tableName = $"{Path.GetFileName(path)}::{sheetName}";
            var jsonPath  = DataIO.GetSchemaJsonPath(path);

            var colNames = DataIO.ReadXlsxHeader(path, sheetName);
            if (colNames.Count == 0) { MessageBox.Show("未能读取到列名。请确认第1行是字段名。"); return; }

            currentWorkbook = File.Exists(jsonPath)
                ? JsonSerializer.Deserialize<WorkbookSchema>(File.ReadAllText(jsonPath), DataIO.JsonOptions)
                : null;

            // 兼容历史：单表 JSON
            if (currentWorkbook == null && File.Exists(jsonPath))
            {
                var legacy = JsonSerializer.Deserialize<TableSchema>(File.ReadAllText(jsonPath), DataIO.JsonOptions);
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

            bool hasSchemaForThisSheet = File.Exists(jsonPath) && entry != null;
            if (!hasSchemaForThisSheet)
            {
                var dataRowsForGuess = DataIO.ReadXlsxRows(path, sheetName, colNames.Count);
                TypeInference.ApplyRecommendedTypes(cols, dataRowsForGuess);
            }

            currentSchema = new TableSchema
            {
                TableName = tableName,
                Columns = cols,
                TableProperties = entry?.SheetProperties ?? new List<PropertyKV>()
            };
            rows.Clear(); foreach (var c in cols) rows.Add(ColumnRow.FromSchema(c));
            Text = $"SchemaStudio — {tableName}";
            RefreshTablePropSummary();
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
        private void RefreshTablePropSummary()
        {
            // 当前 schema 的表属性摘要
            var kvs = currentSchema?.TableProperties ?? new List<PropertyKV>();
            string sum = string.Join("; ", kvs
                .Where(p => !string.IsNullOrWhiteSpace(p.PropertyName))
                .Select(p => $"{p.PropertyName}={p.PropertyValue}"));

            lblTablePropSummary.Text = string.IsNullOrEmpty(sum) ? "表属性：—" : $"表属性：{sum}";
        }

        // ======= 保存 ======= //
        private void SaveSchema()
        {
            if (currentTag == null || currentSchema == null) return;

            foreach (var r in rows)
            {
                var c = currentSchema.Columns!.First(x => x.ColName == r.ColName);
                c.ValueType = r.ValueType;
                c.PkOrder   = r.PkOrder;
                c.PrimaryKey = r.PkOrder >= 0;
            }

            var jsonPath = DataIO.GetSchemaJsonPath(currentTag.Path);

            if (currentTag.IsCsv)
            {
                var json = JsonSerializer.Serialize(currentSchema, DataIO.JsonOptions);
                File.WriteAllText(jsonPath, json, new UTF8Encoding(false));
            }
            else if (currentTag.IsExcelSheet)
            {
                currentWorkbook ??= new WorkbookSchema { WorkbookName = Path.GetFileName(currentTag.Path), Sheets = new() };
                var entry = currentWorkbook.Sheets.FirstOrDefault(s => s.SheetName.Equals(currentTag.Sheet!, StringComparison.OrdinalIgnoreCase));
                if (entry == null) { entry = new SheetSchema { SheetName = currentTag.Sheet!, Columns = new() }; currentWorkbook.Sheets.Add(entry); }
                entry.Columns = currentSchema.Columns!;
                entry.SheetProperties = currentSchema.TableProperties ?? new List<PropertyKV>();
                var json = JsonSerializer.Serialize(currentWorkbook, DataIO.JsonOptions);
                File.WriteAllText(jsonPath, json, new UTF8Encoding(false));
            }
            else if (currentTag.IsExcelFile)
            {
                MessageBox.Show("请选择具体工作表后再保存。");
                return;
            }

            MessageBox.Show($"已保存：{jsonPath}", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // ======= 校验 ======= //
        private void ValidateCurrent()
        {
            if (currentTag == null || currentSchema == null)
            { MessageBox.Show("请先选择 CSV 或 Excel 的某个工作表。"); return; }

            var headers = currentTag.IsExcelSheet ? DataIO.ReadXlsxHeader(currentTag.Path, currentTag.Sheet!) : DataIO.ReadCsvHeader(currentTag.Path);
            if (headers.Count == 0) { MessageBox.Show("未能读取到列名。"); return; }
            var data = currentTag.IsExcelSheet ? DataIO.ReadXlsxRows(currentTag.Path, currentTag.Sheet!, headers.Count) : DataIO.ReadCsvRows(currentTag.Path, headers.Count);

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
                    if (!TypeInference.IsValid(val, col.ValueType ?? "string", out var reason))
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

        // ======= 其它 ======= //
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
        private void CreateNewTable()
        {
            if (currentDir == null) { MessageBox.Show("请先打开根目录"); return; }
            var dirs = DataIO.GetAllDirs(currentDir).ToList();
            using var dlg = new NewTableForm(currentDir, dirs, registry, UiDefaults.AllowedTypes);
            if (dlg.ShowDialog(this) != DialogResult.OK) return;

            var destDir  = dlg.ResultDir!;
            var fileBase = dlg.ResultFileName!;
            var fmt      = dlg.ResultFormat;
            var perSheet = dlg.ResultSheetColumns; 
            var perSheetProps = dlg.ResultSheetTableProps;

            string path = Path.Combine(destDir, fileBase + (fmt == "xlsx" ? ".xlsx" : fmt == "xls" ? ".xls" : ".csv"));
            if (File.Exists(path)) { MessageBox.Show("同名文件已存在，已取消。"); return; }
            Directory.CreateDirectory(destDir);

            bool wrote = false; string warn = "";

            try
            {
                if (fmt == "xlsx")
                    wrote = DataIO.TryWriteXlsxMulti(path, perSheet.Select(kv => (kv.Key, kv.Value.Select(c => c.ColName).ToArray())), out warn);
                else if (fmt == "xls")
                    wrote = DataIO.TryWriteXlsMulti(path,  perSheet.Select(kv => (kv.Key, kv.Value.Select(c => c.ColName).ToArray())), out warn);
            }
            catch (Exception ex) { warn = $"写入 Excel 失败：{ex.Message}"; wrote = false; }

            if (!wrote)
            {
                path = Path.Combine(destDir, fileBase + ".csv");
                var headers = (perSheet.Count > 0 ? perSheet.First().Value : new List<NewColRow>())
                                .Select(c => c.ColName).ToArray();
                DataIO.WriteCsvHeader(path, headers);
                if (!string.IsNullOrEmpty(warn)) MessageBox.Show(warn + "\n已回退为 CSV。");
            }

            // 生成 schema JSON
            var jsonPath = DataIO.GetSchemaJsonPath(path);
            if (fmt == "csv" || !wrote)
            {
                var firstSheetName = perSheet.Keys.FirstOrDefault() ?? "Sheet1";
                var cols = (perSheet.Count > 0 ? perSheet.First().Value : new List<NewColRow>()).ToList();
                var propsForCsv = perSheetProps.TryGetValue(firstSheetName, out var p) ? p : new List<PropertyKV>();
                var schema = new TableSchema
                {
                    TableName = Path.GetFileName(path),
                    Columns = cols.Select((c, idx) => new ColumnSchema
                    {
                        ColName = c.ColName, ValueType = c.ValueType, PkOrder = c.PkOrder,
                        PrimaryKey = c.PkOrder >= 0, ColSort = idx, ColProperties = c.ColProperties?.ToList() ?? new()
                    }).ToList(),
                    TableProperties = propsForCsv
                };
                var json = JsonSerializer.Serialize(schema, DataIO.JsonOptions);
                File.WriteAllText(jsonPath, json, new UTF8Encoding(false));
            }
            else
            {
                var wb = new WorkbookSchema
                {
                    WorkbookName = Path.GetFileName(path),
                    Sheets = perSheet.Select(kv => new SheetSchema
                    {
                        SheetName = kv.Key,
                        Columns = kv.Value.Select((c, idx) => new ColumnSchema
                        {
                            ColName = c.ColName, ValueType = c.ValueType, PkOrder = c.PkOrder,
                            PrimaryKey = c.PkOrder >= 0, ColSort = idx, ColProperties = c.ColProperties?.ToList() ?? new()
                        }).ToList(),
                        SheetProperties = perSheetProps.TryGetValue(kv.Key, out var tp) ? tp : new List<PropertyKV>()
                    }).ToList()
                };
                var json = JsonSerializer.Serialize(wb, DataIO.JsonOptions);
                File.WriteAllText(jsonPath, json, new UTF8Encoding(false));
            }

            BuildDirectoryTree();
            SelectPathInTree(path);
            MessageBox.Show($"已创建：\n{path}\n{path}.schema.json");
        }
        private void SelectPathInTree(string fullPath)
        {
            foreach (TreeNode root in tree.Nodes)
            {
                var node = FindNode(root, fullPath);
                if (node != null) { tree.SelectedNode = node; node.EnsureVisible(); break; }
            }
            TreeNode? FindNode(TreeNode n, string path)
            {
                if (n.Tag is NodeTag t && t.Path == path && t.Sheet == null) return n;
                foreach (TreeNode c in n.Nodes) { var got = FindNode(c, path); if (got != null) return got; }
                return null;
            }
        }
    }
}
