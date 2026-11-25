using System.Drawing.Drawing2D;
using SchemaStudio.Models;

namespace SchemaStudio.UI
{
    public class AssignPropertyForm : Form
    {
        private readonly ListBox list;
        private readonly Panel editor;
        private readonly List<PropertyDef> defs;
        private readonly Dictionary<string, PropertyKV> map;
        private PropertyDef? currentDef;
        private Control? currentInput;
        private CheckBox? include;

        public List<PropertyKV> Result => map.Values.ToList();

        public AssignPropertyForm(List<PropertyDef> definitions, List<PropertyKV> existing, string? targetName = null, bool isTableTarget = true)
        {
            Text = string.IsNullOrWhiteSpace(targetName) ? "设置属性" : $"{(isTableTarget ? "表属性" : "列属性")} - {targetName}";
            ClientSize = new Size(680, 480);
            MinimumSize = new Size(560, 380);
            StartPosition = FormStartPosition.CenterParent;
            Font = StudioTheme.UiFont(9F);
            AutoScaleMode = AutoScaleMode.Dpi;
            BackColor = StudioTheme.Surface;
            ForeColor = StudioTheme.Text;
            HandleCreated += (_, _) => StudioTheme.ApplyModernTitleBar(this);

            defs = definitions;
            map = new Dictionary<string, PropertyKV>(StringComparer.OrdinalIgnoreCase);
            foreach (var p in existing)
                if (!string.IsNullOrWhiteSpace(p.PropertyName))
                    map[p.PropertyName] = new PropertyKV(p.PropertyName, p.PropertyValue ?? "");

            StudioCard? targetBar = null;
            if (!string.IsNullOrWhiteSpace(targetName))
            {
                targetBar = new StudioCard { Dock = DockStyle.Top, BorderAll = false, BorderBottom = true, BackColor = StudioTheme.Background, Padding = new Padding(14, 8, 14, 8) };
                var row = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false, BackColor = StudioTheme.Background, Margin = Padding.Empty, Padding = Padding.Empty };
                var iconBox = new Panel { Size = new Size(16, 22), Margin = new Padding(0, 1, 6, 0), BackColor = StudioTheme.Background };
                bool isCsv = targetName.EndsWith(".csv", StringComparison.OrdinalIgnoreCase);
                iconBox.Paint += (_, e) => StudioTheme.DrawIcon(e.Graphics, isCsv ? StudioIcon.Csv : StudioIcon.Table, new Rectangle(0, 3, 15, 15), isCsv ? StudioTheme.CsvBlue : StudioTheme.SheetGreen);
                var prefix = new Label { Text = isTableTarget ? "当前配置表：" : "当前配置列：", AutoSize = true, ForeColor = StudioTheme.Muted, Margin = new Padding(0, 2, 4, 0) };
                var nameLbl = new Label { Text = targetName, AutoSize = true, ForeColor = StudioTheme.Text, Font = StudioTheme.UiFont(9F, FontStyle.Bold), Margin = new Padding(0, 2, 0, 0) };
                row.Controls.Add(iconBox);
                row.Controls.Add(prefix);
                row.Controls.Add(nameLbl);
                targetBar.Height = StudioTheme.LineHeight(Font) + 18;
                targetBar.Controls.Add(row);
            }

            var body = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Padding = new Padding(14) };
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 250));
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            var leftWrap = new StudioCard { Dock = DockStyle.Fill, Padding = new Padding(1), BackColor = StudioTheme.Background };
            list = new ListBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                IntegralHeight = false,
                DrawMode = DrawMode.OwnerDrawFixed,
                BackColor = StudioTheme.Background,
                ForeColor = StudioTheme.Text
            };
            UpdateListMetrics();
            list.DrawItem += DrawListItem;
            list.SelectedIndexChanged += (_, _) => { CommitEditor(); RenderEditor(); };
            Load += (_, _) => { UpdateListMetrics(); list.Invalidate(); };
            DpiChanged += (_, _) => { UpdateListMetrics(); list.Invalidate(); };
            leftWrap.Controls.Add(list);
            body.Controls.Add(leftWrap, 0, 0);

            editor = new StudioCard { Dock = DockStyle.Fill, BackColor = StudioTheme.Surface, Padding = new Padding(18, 14, 18, 14), Margin = new Padding(12, 0, 0, 0) };
            body.Controls.Add(editor, 1, 0);

            var footer = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(14, 10, 14, 12), BackColor = StudioTheme.Background };
            footer.Resize += (_, _) => footer.Invalidate();
            footer.Paint += (_, e) => { using var p = new Pen(StudioTheme.Border); e.Graphics.DrawLine(p, 0, 0, footer.Width, 0); };

            var btnOk = StudioTheme.Button("确定", (_, _) =>
            {
                CommitEditor();
                var missing = defs.FirstOrDefault(d => d.Kind == "enum" && map.TryGetValue(d.Key, out var property) && string.IsNullOrEmpty(property.PropertyValue));
                if (missing != null) { list.SelectedItem = missing; MessageBox.Show(this, "请选择属性：" + missing.Display); return; }
                DialogResult = DialogResult.OK; Close();
            }, primary: true);
            var btnCancel = StudioTheme.Button("取消", (_, _) => { DialogResult = DialogResult.Cancel; Close(); });
            footer.Controls.Add(btnOk);
            footer.Controls.Add(btnCancel);

            Controls.Add(body);
            if (targetBar != null) Controls.Add(targetBar);
            Controls.Add(footer);
            AcceptButton = btnOk;
            CancelButton = btnCancel;

            foreach (var d in defs) list.Items.Add(d);
            if (list.Items.Count > 0) list.SelectedIndex = 0;
            else editor.Controls.Add(new Label { Text = "当前没有可用属性，请先在「属性库」中添加对应适用范围的属性。", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, ForeColor = StudioTheme.Muted });
        }

        private int titleH = 20, subH = 18;

        private void UpdateListMetrics()
        {
            using var titleFont = StudioTheme.UiFont(9F, FontStyle.Bold);
            using var subFont = StudioTheme.UiFont(8.25F);
            titleH = StudioTheme.LineHeight(titleFont);
            subH = StudioTheme.LineHeight(subFont);
            list.ItemHeight = titleH + subH + 18;
        }

        private void DrawListItem(object? sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= list.Items.Count) return;
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (var bg = new SolidBrush(StudioTheme.Background)) g.FillRectangle(bg, e.Bounds);

            bool selected = (e.State & DrawItemState.Selected) != 0;
            if (selected)
            {
                var pill = new Rectangle(e.Bounds.X + 6, e.Bounds.Y + 3, e.Bounds.Width - 12, e.Bounds.Height - 6);
                using var path = StudioTheme.RoundedRect(pill, 6);
                using var fill = new SolidBrush(StudioTheme.ActiveTab);
                g.FillPath(fill, path);
            }

            if (list.Items[e.Index] is not PropertyDef d) return;
            bool enabled = map.TryGetValue(d.Key, out var kv);
            string title = string.IsNullOrWhiteSpace(d.Title) ? d.Key : d.Title;
            string sub = enabled ? $"已启用 · {kv!.PropertyValue}" : $"{d.Key} · 未启用";

            using var titleFont = StudioTheme.UiFont(9F, selected ? FontStyle.Bold : FontStyle.Regular);
            using var subFont = StudioTheme.UiFont(8.25F);
            int top = e.Bounds.Y + Math.Max(4, (e.Bounds.Height - (titleH + subH + 2)) / 2);
            TextRenderer.DrawText(g, title, titleFont,
                new Rectangle(e.Bounds.X + 14, top, e.Bounds.Width - 28, titleH),
                StudioTheme.Text, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
            TextRenderer.DrawText(g, sub, subFont,
                new Rectangle(e.Bounds.X + 14, top + titleH + 2, e.Bounds.Width - 28, subH),
                enabled ? StudioTheme.Accent : StudioTheme.Muted, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
        }

        private void RenderEditor()
        {
            editor.Controls.Clear();
            currentInput = null;
            currentDef = list.SelectedItem as PropertyDef;
            if (currentDef == null) return;

            var d = currentDef;
            var stack = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1 };
            stack.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            stack.Controls.Add(new Label
            {
                Text = string.IsNullOrEmpty(d.Title) ? d.Key : $"{d.Title} ({d.Key})",
                AutoSize = true,
                Font = StudioTheme.UiFont(10F, FontStyle.Bold),
                ForeColor = StudioTheme.Text,
                Margin = new Padding(0, 0, 0, 12)
            });

            map.TryGetValue(d.Key, out var kv);
            var v = kv?.PropertyValue ?? d.Default ?? string.Empty;

            include = new CheckBox { Text = "启用此属性", Checked = map.ContainsKey(d.Key), AutoSize = true, Margin = new Padding(0, 0, 0, 10) };
            stack.Controls.Add(include);

            switch (d.Kind)
            {
                case "bool":
                    var chk = new CheckBox { Text = "true / false", Checked = v.Equals("true", StringComparison.OrdinalIgnoreCase), AutoSize = true };
                    chk.CheckedChanged += (_, _) => { CommitEditor(); list.Invalidate(); };
                    currentInput = chk;
                    break;
                case "int":
                    var num = new NumericUpDown { Width = 160, Minimum = int.MinValue, Maximum = int.MaxValue, Value = int.TryParse(v, out var iv) ? iv : 0, BorderStyle = BorderStyle.FixedSingle };
                    num.ValueChanged += (_, _) => { CommitEditor(); list.Invalidate(); };
                    currentInput = num;
                    break;
                case "string":
                    var tb = new TextBox { Dock = DockStyle.Top, Text = v, BorderStyle = BorderStyle.FixedSingle };
                    tb.TextChanged += (_, _) => { CommitEditor(); list.Invalidate(); };
                    currentInput = tb;
                    break;
                default:
                    var cb = new ComboBox { Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDownList };
                    cb.Items.AddRange((d.EnumValues ?? new List<string>()).ToArray());
                    if (!string.IsNullOrEmpty(v) && !cb.Items.Contains(v)) cb.Items.Add(v);
                    cb.SelectedItem = string.IsNullOrEmpty(v) ? null : v;
                    cb.SelectedIndexChanged += (_, _) => { CommitEditor(); list.Invalidate(); };
                    currentInput = cb;
                    break;
            }
            currentInput.Enabled = include.Checked;
            include.CheckedChanged += (_, _) =>
            {
                if (currentInput != null) currentInput.Enabled = include.Checked;
                CommitEditor();
                list.Invalidate();
            };
            stack.Controls.Add(currentInput);
            editor.Controls.Add(stack);
        }

        private void CommitEditor()
        {
            if (currentDef == null || currentInput == null) return;
            var d = currentDef;
            if (include?.Checked != true) { map.Remove(d.Key); return; }
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
}
