namespace SchemaStudio.Services
{
    internal static class UiDefaults
    {
        public const int LeftTreeWidth = 35;
        public static readonly Dictionary<string, int> GridColWidths = new()
        {
            ["col_name"]     = 260,
            ["value_type"]   = 140,
            ["pk_order"]     = 150,
            ["prop_btn"]     = 80,
            ["prop_summary"] = 360
        };

        public static readonly string[] FileExts = new[] { ".csv", ".xlsx", ".xls" };
        public static readonly string[] AllowedTypes =
        {
            "int","long","float","double","bool","string","int[]","List<int>"
        };
    }
}
