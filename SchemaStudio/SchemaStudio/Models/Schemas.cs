using Newtonsoft.Json;
using System.ComponentModel;

namespace SchemaStudio.Models
{
    // —— 节点标记：文件 / Sheet 两级 —— //
    public record NodeTag(string Path, string? Sheet)
    {
        public bool IsExcelFile  => (Path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || Path.EndsWith(".xls", StringComparison.OrdinalIgnoreCase)) && Sheet == null;
        public bool IsExcelSheet => Sheet != null && (Path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || Path.EndsWith(".xls", StringComparison.OrdinalIgnoreCase));
        public bool IsCsv        => Path.EndsWith(".csv", StringComparison.OrdinalIgnoreCase);
        public override string ToString() => Sheet == null ? Path : $"{Path}::{Sheet}";
    }

    // —— 新建表对话框用的行模型 —— //
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
                else if (int.TryParse(value.Trim(), out var v)) PkOrder = v; else PkOrder = -1;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(PkOrderText)));
            }
        }
    }
    public class SheetNameRow { public string Name { get; set; } = "Sheet1"; }

    // —— 主界面右侧表格绑定行 —— //
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
                else if (int.TryParse(value.Trim(), out var v)) PkOrder = v; else PkOrder = -1;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(PkOrderText)));
            }
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

    // —— JSON Schema —— //
    public class TableSchema
    {
        [JsonProperty("table_name")]       public string? TableName { get; set; }
        [JsonProperty("columns")]          public List<ColumnSchema>? Columns { get; set; }
        [JsonProperty("table_properties")] public List<PropertyKV>? TableProperties { get; set; } = new();
    }
    public class WorkbookSchema
    {
        [JsonProperty("workbook_name")] public string? WorkbookName { get; set; }
        [JsonProperty("sheets")]        public List<SheetSchema> Sheets { get; set; } = new();
    }
    public class SheetSchema
    {
        [JsonProperty("sheet_name")]       public string SheetName { get; set; } = "";
        [JsonProperty("columns")]          public List<ColumnSchema> Columns { get; set; } = new();
        [JsonProperty("sheet_properties")] public List<PropertyKV>? SheetProperties { get; set; } = new();
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

    // —— 属性库 —— //
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
        [JsonProperty("kind")]    public string Kind { get; set; } = "enum";
        [JsonProperty("values")]  public List<string>? EnumValues { get; set; } = new();
        [JsonProperty("default")] public string? Default { get; set; }
        [JsonIgnore] public string Display => string.IsNullOrEmpty(Title) ? Key : $"{Key}（{Title}）";
        public override string ToString() => Display;
    }

    public class ValidationError
    {
        public int Row { get; set; }
        public string Column { get; set; } = "";
        public string Expected { get; set; } = "";
        public string Value { get; set; } = "";
        public string Message { get; set; } = "";
    }
}
