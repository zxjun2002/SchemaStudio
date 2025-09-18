using SchemaStudio.Models;

namespace SchemaStudio.UI
{
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
                default: // enum
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
}
