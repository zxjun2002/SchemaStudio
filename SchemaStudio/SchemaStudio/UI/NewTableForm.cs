using System.ComponentModel;
using SchemaStudio.Models;

namespace SchemaStudio.UI
{
    public class NewTableForm : Form
    {
        private readonly ComboBox cmbDir, cmbFormat;
        private readonly TextBox txtName;
        private readonly DataGridView gridCols, gridSheets;
        private readonly Button btnAddCol, btnDelCol, btnAddSheet, btnDelSheet, btnOk, btnCancel;
        private readonly Button btnTableProps;
        private readonly Label  lblTablePropSummary;

        private readonly BindingList<SheetNameRow> sheets = new();
        private readonly Dictionary<string, BindingList<NewColRow>> sheetCols =
            new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, List<PropertyKV>> sheetTableProps =
            new(StringComparer.OrdinalIgnoreCase);

        private string CurrentSheet => GetSelectedSheetNameOrFirst();

        private readonly PropertyRegistry registry;
        private readonly string root;
        private readonly string[] allowed;

        public string? ResultDir { get; private set; }
        public string? ResultFileName { get; private set; }
        public string ResultFormat { get; private set; } = "csv";
        //正在编辑的Sheet名
        private string? _oldSheetNameEditing;

        public List<string> ResultSheets => sheets.Select(s => s.Name).Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
        public Dictionary<string, List<NewColRow>> ResultSheetColumns =>
            sheetCols.ToDictionary(kv => kv.Key, kv => kv.Value.ToList(), StringComparer.OrdinalIgnoreCase);
        public IEnumerable<NewColRow> Columns =>
            sheetCols.TryGetValue(CurrentSheet, out var list) ? list : Enumerable.Empty<NewColRow>();
        
        public Dictionary<string, List<PropertyKV>> ResultSheetTableProps =>
            sheetTableProps.ToDictionary(kv => kv.Key, kv => kv.Value?.ToList() ?? new List<PropertyKV>(),
                StringComparer.OrdinalIgnoreCase);

        public NewTableForm(string root, List<string> allDirs, PropertyRegistry reg, string[] allowedTypes)
        {
            this.root = root; registry = reg; allowed = allowedTypes;

            Text = "新建表"; Width = 980; Height = 640; AutoScaleMode = AutoScaleMode.Font;

            var lblDir = new Label { Left = 12, Top = 14, Text = "保存到目录：", AutoSize = true };
            cmbDir = new ComboBox
            {
                Left = 90, Top = 10, Width = 680, DropDownStyle = ComboBoxStyle.DropDownList,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            cmbDir.Items.AddRange(allDirs
                .Select(d => d.Replace(root, "").TrimStart(Path.DirectorySeparatorChar))
                .Select(s => string.IsNullOrEmpty(s) ? "(根目录)" : s)
                .ToArray<object>());
            cmbDir.SelectedIndex = 0;

            var btnBrowse = new Button { Left = 780, Top = 10, Width = 70, Text = "浏览..." };
            btnBrowse.Click += (_, __) =>
            {
                using var fbd = new FolderBrowserDialog { SelectedPath = root };
                if (fbd.ShowDialog(this) == DialogResult.OK)
                {
                    var p = fbd.SelectedPath;
                    var show = p.StartsWith(root, StringComparison.OrdinalIgnoreCase)
                        ? p.Replace(root, "").TrimStart(Path.DirectorySeparatorChar)
                        : p;
                    if (string.IsNullOrEmpty(show)) show = "(根目录)";
                    if (!cmbDir.Items.Cast<object>().Any(x => string.Equals(x.ToString(), show, StringComparison.OrdinalIgnoreCase)))
                        cmbDir.Items.Add(show);
                    cmbDir.SelectedItem = show;
                }
            };

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

            gridCols = new DataGridView
            {
                Left = 12, Top = 80, Width = 940, Height = 380,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                AutoGenerateColumns = false, AllowUserToAddRows = false
            };
            gridCols.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "列名", DataPropertyName = nameof(NewColRow.ColName), Width = 240 });
            gridCols.Columns.Add(new DataGridViewComboBoxColumn { HeaderText = "数值类型", DataPropertyName = nameof(NewColRow.ValueType), Width = 140, DataSource = allowed });
            gridCols.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "主键序列", DataPropertyName = nameof(NewColRow.PkOrderText), Width = 120 });
            var btnColProp = new DataGridViewButtonColumn { HeaderText = "列属性", Text = "属性…", UseColumnTextForButtonValue = true, Width = 80 };
            gridCols.Columns.Add(btnColProp);
            gridCols.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "属性摘要", DataPropertyName = nameof(NewColRow.PropsSummary), AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, ReadOnly = true });
            gridCols.CellContentClick += (s, e) =>
            {
                if (e.RowIndex < 0) return;
                if (gridCols.Columns[e.ColumnIndex] is DataGridViewButtonColumn &&
                    gridCols.Rows[e.RowIndex].DataBoundItem is NewColRow row)
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
            btnAddCol.Click += (_, __) =>
            {
                var list = GetCurrentCols();
                list.Add(new NewColRow { ColName = $"col_{list.Count + 1}", ValueType = "string", PkOrder = -1 });
            };
            btnDelCol.Click += (_, __) =>
            {
                var list = GetCurrentCols();
                foreach (DataGridViewRow r in gridCols.SelectedRows)
                    if (r.DataBoundItem is NewColRow x) list.Remove(x);
            };
            
            btnTableProps = new Button { Text = "表属性…", Left = 200, Top = 470, Width = 90, Anchor = AnchorStyles.Bottom | AnchorStyles.Left };
            lblTablePropSummary = new Label { Text = "表属性：—", Left = 300, Top = 475, AutoSize = true, Anchor = AnchorStyles.Bottom | AnchorStyles.Left };

            btnTableProps.Click += (_, __) =>
            {
                var defs = registry.Properties.Where(d => d.Scope == "column" || d.Scope == "both").ToList();
                var key = CurrentSheet;
                if (!sheetTableProps.TryGetValue(key, out var list))
                    sheetTableProps[key] = list = new List<PropertyKV>();

                using var dlg = new AssignPropertyForm(defs, list);
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    sheetTableProps[key] = dlg.Result ?? new List<PropertyKV>();
                    RefreshTablePropSummaryFor(key);
                }
            };

            // 把控件加入窗体
            Controls.Add(btnTableProps);
            Controls.Add(lblTablePropSummary);

            gridSheets = new DataGridView
            {
                Left = 12, Top = 510, Width = 480, Height = 80,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left,
                AutoGenerateColumns = false, AllowUserToAddRows = false,
                Visible = false, Enabled = false
            };
            gridSheets.DataSource = sheets;
            gridSheets.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Sheet 名", DataPropertyName = nameof(SheetNameRow.Name), Width = 360 });

            gridSheets.CellBeginEdit += (_, e) =>
            {
                if (e.RowIndex >= 0 && gridSheets.Rows[e.RowIndex].DataBoundItem is SheetNameRow s)
                    _oldSheetNameEditing = s.Name;   // 记住旧名
            };

            gridSheets.SelectionChanged += (_, __) =>
            {
                if (gridSheets.IsCurrentCellInEditMode) return; // 编辑过程中别切数据源
                BindColsToCurrentSheet();
                RefreshTablePropSummaryFor(CurrentSheet);
            };
            
            gridSheets.CellEndEdit += (_, e) =>
            {
                if (e.RowIndex < 0) return;

                var s = (SheetNameRow)gridSheets.Rows[e.RowIndex].DataBoundItem!;
                var newName = (s.Name ?? "").Trim();
                if (string.IsNullOrWhiteSpace(newName))
                {
                    MessageBox.Show("Sheet 名不能为空");
                    s.Name = gridSheets.Rows[e.RowIndex].Tag as string ?? "Sheet1";
                    gridSheets.InvalidateRow(e.RowIndex);
                    return;
                }

                var oldName = _oldSheetNameEditing ?? (gridSheets.Rows[e.RowIndex].Tag as string ?? newName);
                _oldSheetNameEditing = null;

                // 重名校验（排除没变化的情况）
                if (!string.Equals(oldName, newName, StringComparison.OrdinalIgnoreCase) &&
                    sheetCols.ContainsKey(newName))
                {
                    MessageBox.Show("Sheet 名重复");
                    s.Name = oldName;
                    gridSheets.InvalidateRow(e.RowIndex);
                    return;
                }

                // 迁移列集合
                if (!sheetCols.TryGetValue(oldName, out var colsList))
                    colsList = new BindingList<NewColRow>();
                sheetCols.Remove(oldName);
                sheetCols[newName] = colsList;

                // 迁移表属性
                if (sheetTableProps.Remove(oldName, out var tp))
                    sheetTableProps[newName] = tp;
                if (!sheetTableProps.ContainsKey(newName))
                    sheetTableProps[newName] = new List<PropertyKV>();

                // 更新 Tag 记录
                gridSheets.Rows[e.RowIndex].Tag = newName;

                // 显式重绑，避免绑定到了“改名过程中临时创建的空列表”
                gridCols.DataSource = sheetCols[newName];
                RefreshTablePropSummaryFor(newName);
            };

            btnAddSheet = new Button { Text = "添加Sheet", Left = 510, Top = 516, Width = 90, Anchor = AnchorStyles.Bottom | AnchorStyles.Left, Visible = false, Enabled = false };
            btnDelSheet = new Button { Text = "删除选中", Left = 610, Top = 516, Width = 90, Anchor = AnchorStyles.Bottom | AnchorStyles.Left, Visible = false, Enabled = false };
            btnAddSheet.Click += (_, __) =>
            {
                var name = AutoSheetName();
                sheets.Add(new SheetNameRow { Name = name });
                sheetCols[name] = new BindingList<NewColRow>
                {
                    new() { ColName = "id",   ValueType = "int",    PkOrder = 1 },
                    new() { ColName = "desc", ValueType = "string", PkOrder = -1 }
                };
                gridSheets.ClearSelection();
                int last = gridSheets.Rows.Count - 1;
                if (last >= 0)
                {
                    gridSheets.Rows[last].Selected = true;
                    gridSheets.CurrentCell = gridSheets.Rows[last].Cells[0];
                    gridSheets.Rows[last].Tag = name;
                }
                // 新 sheet 的表属性为空集合
                if (!sheetTableProps.ContainsKey(name)) sheetTableProps[name] = new List<PropertyKV>();
                BindColsToCurrentSheet();
                RefreshTablePropSummaryFor(name);
            };
            btnDelSheet.Click += (_, __) =>
            {
                if (gridSheets.CurrentRow?.DataBoundItem is not SheetNameRow s) return;
                if (sheets.Count <= 1) { MessageBox.Show("至少保留一个 Sheet"); return; }
                sheetCols.Remove(s.Name);
                sheetTableProps.Remove(s.Name);
                sheets.Remove(s);
                BindColsToCurrentSheet();
                if (gridSheets.CurrentRow?.DataBoundItem is SheetNameRow cur)
                    RefreshTablePropSummaryFor(cur.Name);
                else
                    lblTablePropSummary.Text = "表属性：—";
            };

            btnOk     = new Button { Text = "确定", Left = 780, Top = 560, Width = 80, Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            btnCancel = new Button { Text = "取消", Left = 872, Top = 560, Width = 80, Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            btnOk.Click += (_, __) => OnOk(allDirs);
            btnCancel.Click += (_, __) => { DialogResult = DialogResult.Cancel; Close(); };

            Controls.AddRange(new Control[] { lblDir, cmbDir, btnBrowse, lblName, txtName, lblFmt, cmbFormat,
                gridCols, btnAddCol, btnDelCol, gridSheets, btnAddSheet, btnDelSheet, btnOk, btnCancel });

            // 默认 1 个 Sheet + 两列
            var first = "Sheet1";
            sheetTableProps[first] = new List<PropertyKV>(); // 新增
            RefreshTablePropSummaryFor(first);               // 新增
            sheets.Add(new SheetNameRow { Name = first });
            sheetCols[first] = new BindingList<NewColRow>
            {
                new() { ColName = "id",   ValueType = "int",    PkOrder = 1 },
                new() { ColName = "desc", ValueType = "string", PkOrder = -1 }
            };
            gridSheets.ClearSelection();
            if (gridSheets.Rows.Count > 0)
            {
                gridSheets.Rows[0].Selected = true;
                gridSheets.CurrentCell = gridSheets.Rows[0].Cells[0];
                gridSheets.Rows[0].Tag = first;
            }
            BindColsToCurrentSheet();
        }

        private BindingList<NewColRow> GetCurrentCols()
        {
            var key = CurrentSheet;
            if (!sheetCols.TryGetValue(key, out var list))
            {
                list = new BindingList<NewColRow>(); sheetCols[key] = list;
            }
            gridCols.DataSource = list;
            return list;
        }
        private void BindColsToCurrentSheet()
        {
            var key = CurrentSheet;
            if (!sheetCols.TryGetValue(key, out var list))
            {
                list = new BindingList<NewColRow>(); sheetCols[key] = list;
            }
            gridCols.DataSource = list;
        }
        private string GetSelectedSheetNameOrFirst()
        {
            if (gridSheets.CurrentRow?.DataBoundItem is SheetNameRow cur && !string.IsNullOrWhiteSpace(cur.Name))
                return cur.Name;
            if (gridSheets.Rows.Count > 0 &&
                gridSheets.Rows[0].DataBoundItem is SheetNameRow first &&
                !string.IsNullOrWhiteSpace(first.Name))
                return first.Name;
            if (sheetCols.Count > 0) return sheetCols.Keys.First();
            return "Sheet1";
        }
        private string AutoSheetName()
        {
            int i = sheets.Count + 1; string name;
            do { name = $"Sheet{i++}"; } while (sheets.Any(s => s.Name.Equals(name, StringComparison.OrdinalIgnoreCase)));
            return name;
        }
        private void OnOk(List<string> allDirs)
        {
            var name = (txtName.Text ?? "").Trim();
            if (string.IsNullOrEmpty(name)) { MessageBox.Show("文件名不能为空"); return; }
            if (name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) { MessageBox.Show("文件名包含非法字符"); return; }

            var rel = cmbDir.SelectedItem!.ToString()!;
            string destDir = rel == "(根目录)" ? root :
                (Path.IsPathRooted(rel) ? rel : Path.Combine(root, rel.TrimStart(Path.DirectorySeparatorChar)));
            if (!Directory.Exists(destDir)) { MessageBox.Show("目标目录无效"); return; }

            ResultFormat = cmbFormat.SelectedItem!.ToString()!;
            if (ResultFormat is "xlsx" or "xls")
            {
                if (sheets.Count == 0) { MessageBox.Show("请至少添加一个工作表"); return; }
                if (sheets.Any(s => string.IsNullOrWhiteSpace(s.Name))) { MessageBox.Show("存在空的工作表名"); return; }
                if (sheets.Select(s => s.Name).Distinct(StringComparer.OrdinalIgnoreCase).Count() != sheets.Count)
                { MessageBox.Show("工作表名重复"); return; }
                foreach (var s in sheets)
                {
                    if (!sheetCols.TryGetValue(s.Name, out var cols) || cols.Count == 0)
                    { MessageBox.Show($"Sheet \"{s.Name}\" 至少需要一列"); return; }
                }
            }
            else
            {
                if (!sheetCols.TryGetValue(CurrentSheet, out var cols) || cols.Count == 0)
                { MessageBox.Show("请至少添加一列"); return; }
            }

            ResultDir = destDir;
            ResultFileName = name;
            DialogResult = DialogResult.OK;
            Close();
        }
        
        private void RefreshTablePropSummaryFor(string sheet)
        {
            if (!sheetTableProps.TryGetValue(sheet, out var list) || list == null || list.Count == 0)
            {
                lblTablePropSummary.Text = "表属性：—";
                return;
            }
            lblTablePropSummary.Text = "表属性：" + string.Join("; ", list.Select(p => $"{p.PropertyName}={p.PropertyValue}"));
        }
    }
}
