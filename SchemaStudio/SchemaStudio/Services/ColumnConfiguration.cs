using SchemaStudio.Models;

namespace SchemaStudio.Services;

public static class ColumnConfiguration
{
    public static List<PropertyDef> Definitions(PropertyRegistry registry) => registry.Properties.Where(d => d.Scope is "column" or "both").ToList();

    public static string? Missing(ColumnSchema column, PropertyRegistry registry)
    {
        return !column.IsConfigured || string.IsNullOrWhiteSpace(column.ValueType) ? "此列缺少类型配置，请选择或确认数据类型。" : null;
    }


    public static void Confirm(EditorDocument document, ColumnSchema column, PropertyRegistry registry)
    {
        var before = column.ColProperties; bool configured = column.IsConfigured;
        var after = before ?? new List<PropertyKV>();
        if (string.IsNullOrWhiteSpace(column.ValueType)) throw new InvalidOperationException("请选择数据类型。");
        if (configured && Missing(column, registry) == null) return;
        document.Execute("确认字段配置", () => { column.ColProperties = after; column.IsConfigured = true; },
            () => { column.ColProperties = before; column.IsConfigured = configured; });
    }
}
