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

        public AssignPropertyForm(List<PropertyDef> definitions, List<PropertyKV> existing)
        {
            Text = "设置属性";
            ClientSize = new Size(680, 460);
            MinimumSize = new Size(560, 380);
            StartPosition = FormStartPosition.CenterParent;
            Font = StudioTheme.UiFont(9F);
            AutoScaleMode = AutoScaleMode.Dpi;
            BackColor = StudioTheme.Surface;
            ForeColor = StudioTheme.Text;
            HandleCreated += (_, _) => StudioTheme.ApplyModernTitleBar(this);

            defs = definitions;
            map = existing.ToDictionary(p => p.PropertyName!, StringComparer.OrdinalIgnoreCase);

            var body = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Padding = new Padding(14) };
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 250));
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            var leftWrap = new Panel { Dock = DockStyle.Fill, Padding = new Padding(1), BackColor = StudioTheme.Border };
            list = new ListBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                IntegralHeight = false,
                DrawMode = DrawMode.OwnerDrawFixed,
                ItemHeight = 46,
                BackColor = StudioTheme.Background,
                ForeColor = StudioTheme.Text
            };
            list.DrawItem += DrawListItem;
            list.SelectedIndexChanged += (_, _) => { CommitEditor(); RenderEditor(); };
            leftWrap.Controls.Add(list);
            body.Controls.Add(leftWrap, 0, 0);

            editor = new Panel { Dock = DockStyle.Fill, BackColor = StudioTheme.Surface, Padding = new Padding(18, 14, 18, 14), Margin = new Padding(12, 0, 0, 0) };
            editor.Paint += (_, e) => { using var p = new Pen(StudioTheme.Border); e.Graphics.DrawRectangle(p, 0, 0, editor.Width - 1, editor.Height - 1); };
            body.Controls.Add(editor, 1, 0);

            var footer = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(14, 10, 14, 12), BackColor = StudioTheme.Background };
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
            Controls.Add(footer);
            AcceptButton = btnOk;
            CancelButton = btnCancel;

            foreach (var d in defs) list.Items.Add(d);
            if (list.Items.Count > 0) list.SelectedIndex = 0;
            else editor.Controls.Add(new Label { Text = "当前没有可用属性，请先在「属性库」中添加对应适用范围的属性。", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, ForeColor = StudioTheme.Muted });
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

            TextRenderer.DrawText(g, title, StudioTheme.UiFont(9F, selected ? FontStyle.Bold : FontStyle.Regular),
                new Rectangle(e.Bounds.X + 14, e.Bounds.Y + 6, e.Bounds.Width - 28, 18),
                StudioTheme.Text, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
            TextRenderer.DrawText(g, sub, StudioTheme.UiFont(8F),
                new Rectangle(e.Bounds.X + 14, e.Bounds.Y + 24, e.Bounds.Width - 28, 16),
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
