using SchemaStudio.Models;
using System.Globalization;
using System.Text.Json;

namespace SchemaStudio.Services
{
    internal static class TypeInference
    {
        public static bool IsValid(string value, string type, out string reason)
        {
            reason = ""; if (string.IsNullOrWhiteSpace(value)) return true;
            var t = (type ?? "string").Trim(); var ic = CultureInfo.InvariantCulture;

            bool ParseBool(string s)
            {
                s = s.Trim().ToLowerInvariant();
                if (bool.TryParse(s, out _)) return true;
                return s is "1" or "0" or "y" or "n" or "yes" or "no" or "是" or "否" or "真" or "假";
            }

            switch (t)
            {
                case "int":    if (!int.TryParse(value, NumberStyles.Integer, ic, out _))   { reason = "不是有效的 int"; return false; } return true;
                case "long":   if (!long.TryParse(value, NumberStyles.Integer, ic, out _))  { reason = "不是有效的 long"; return false; } return true;
                case "float":  if (!float.TryParse(value, NumberStyles.Float, ic, out _))   { reason = "不是有效的 float"; return false; } return true;
                case "double": if (!double.TryParse(value, NumberStyles.Float, ic, out _))  { reason = "不是有效的 double"; return false; } return true;
                case "bool":   if (!ParseBool(value)) { reason = "不是有效的 bool(true/false/1/0/是/否)"; return false; } return true;
                case "string": return true;
                case "int[]":
                case "List<int>":
                    return TryParseIntArray(value, out _, out reason);
                default:       return true;
            }
        }

        public static bool TryParseIntArray(string s, out int[] result, out string reason)
        {
            result = Array.Empty<int>();
            reason = "";
            if (string.IsNullOrWhiteSpace(s)) return true;

            s = s.Trim();
            if (!(s.StartsWith("[") && s.EndsWith("]")))
            {
                reason = "必须是 JSON 数组格式（示例：[1,2,3]）";
                return false;
            }

            try
            {
                var list = JsonSerializer.Deserialize<List<int>>(s);
                if (list == null) { reason = "不是有效的 JSON 数组（示例：[1,2,3]）"; return false; }
                result = list.ToArray();
                return true;
            }
            catch
            {
                reason = "不是有效的 JSON 数组（示例：[1,2,3]）";
                return false;
            }
        }

        // —— 推荐 —— //
        private static bool TryParseBoolLoose(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return true;
            s = s.Trim().ToLowerInvariant();
            if (bool.TryParse(s, out _)) return true;
            return s is "1" or "0" or "y" or "n" or "yes" or "no" or "是" or "否" or "真" or "假";
        }
        private static bool IsAllInt(IEnumerable<string> cells)
        {
            var ic = CultureInfo.InvariantCulture;
            foreach (var v in cells)
            {
                if (string.IsNullOrWhiteSpace(v)) continue;
                if (!int.TryParse(v, NumberStyles.Integer, ic, out _)) return false;
            }
            return true;
        }
        private static bool IsAllLong(IEnumerable<string> cells)
        {
            var ic = CultureInfo.InvariantCulture;
            foreach (var v in cells)
            {
                if (string.IsNullOrWhiteSpace(v)) continue;
                if (!long.TryParse(v, NumberStyles.Integer, ic, out _)) return false;
            }
            return true;
        }
        private static bool IsAllFloat(IEnumerable<string> cells)
        {
            var ic = CultureInfo.InvariantCulture; bool any = false;
            foreach (var v in cells)
            {
                if (string.IsNullOrWhiteSpace(v)) continue;
                any = true;
                if (!float.TryParse(v, NumberStyles.Float, ic, out _)) return false;
            }
            return any;
        }
        private static bool IsAllDouble(IEnumerable<string> cells)
        {
            var ic = CultureInfo.InvariantCulture; bool any = false;
            foreach (var v in cells)
            {
                if (string.IsNullOrWhiteSpace(v)) continue;
                any = true;
                if (!double.TryParse(v, NumberStyles.Float, ic, out _)) return false;
            }
            return any;
        }
        private static bool IsAllBool(IEnumerable<string> cells)
        {
            foreach (var v in cells)
            {
                if (string.IsNullOrWhiteSpace(v)) continue;
                if (!TryParseBoolLoose(v)) return false;
            }
            return true;
        }
        private static bool IsAllIntArray(IEnumerable<string> cells, out bool sawAny)
        {
            sawAny = false;
            foreach (var v in cells)
            {
                if (string.IsNullOrWhiteSpace(v)) continue;
                sawAny = true;
                if (!TryParseIntArray(v, out _, out _)) return false;
            }
            return sawAny;
        }

        public static string RecommendTypeFromSamples(IEnumerable<string> samples)
        {
            var list = samples.ToList();
            if (IsAllIntArray(list, out var anyArr) && anyArr) return "int[]";
            if (IsAllBool(list)) return "bool";
            if (IsAllInt(list)) return "int";
            if (IsAllLong(list)) return "long";
            if (IsAllFloat(list)) return "float";
            if (IsAllDouble(list)) return "double";
            return "string";
        }

        public static void ApplyRecommendedTypes(List<ColumnSchema> cols, List<string[]> dataRows)
        {
            if (cols == null || cols.Count == 0) return;
            const int N = 120;
            var samples = dataRows.Take(N).ToList();

            for (int ci = 0; ci < cols.Count; ci++)
            {
                var col = cols[ci];
                var cur = (col.ValueType ?? "").Trim();
                if (!string.IsNullOrEmpty(cur) && !string.Equals(cur, "string", StringComparison.OrdinalIgnoreCase))
                    continue;

                var colSamples = samples.Select(r => ci < r.Length ? (r[ci] ?? "") : "");
                var rec = RecommendTypeFromSamples(colSamples);

                if (!string.Equals(rec, "string", StringComparison.OrdinalIgnoreCase))
                    col.ValueType = rec;
                else
                    col.ValueType = string.IsNullOrEmpty(cur) ? "string" : cur;
            }
        }
    }
}
