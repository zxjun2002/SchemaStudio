using System.Text.Json;
using SchemaStudio.Models;
using SchemaStudio.Services;

namespace SchemaStudio.UI
{
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
            // 用 System.Text.Json 做深拷贝，保持与全局 JsonOptions 一致
            working = JsonSerializer.Deserialize<PropertyRegistry>(
                          JsonSerializer.Serialize(reg, DataIO.JsonOptions),
                          DataIO.JsonOptions) ?? new PropertyRegistry();

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
            btnDel.Click += (_, __) =>
            {
                if (list.SelectedItem is PropertyDef d)
                {
                    working.Properties.RemoveAll(p => p.Key.Equals(d.Key, StringComparison.OrdinalIgnoreCase));
                    RefreshList();
                }
            };
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
            var key = txtKey.Text.Trim();
            if (string.IsNullOrEmpty(key)) { MessageBox.Show("key 不能为空"); return; }

            var def = working.Properties.FirstOrDefault(p => p.Key.Equals(key, StringComparison.OrdinalIgnoreCase));
            if (def == null)
            {
                def = new PropertyDef { Key = key, Scope = "column", Kind = "enum" };
                working.Properties.Add(def);
            }

            def.Title = txtTitle.Text.Trim();
            def.Scope = "column";
            def.Kind  = string.IsNullOrEmpty(def.Kind) ? "enum" : def.Kind;
            def.EnumValues = txtEnum.Text
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(s => s.Trim())
                .Distinct()
                .ToList();
            def.Default = txtDefault.Text.Trim();

            RefreshList();
        }
    }
}
