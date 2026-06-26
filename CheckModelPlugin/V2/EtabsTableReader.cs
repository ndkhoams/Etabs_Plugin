using ETABSv1;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;

namespace CheckModelPlugin
{
    public static class EtabsTableReader
    {
        // ─── Tên bảng ETABS phổ biến theo từng loại dữ liệu ──────────────────────
        // Tập trung tại đây thay vì rải rác ở các Extractor.

        internal static readonly string[] MassSummaryTableNames =
        {
            "Mass Summary by Story",
            "Story Mass Summary",
            "Masses by Story",
            "Center Of Mass And Rigidity",
            "Centers Of Mass And Rigidity"
        };

        internal static readonly string[] DiaphragmDisplacementTableNames =
        {
            "Diaphragm Center of Mass Displacements",
            "Diaphragm Center Of Mass Displacements",
            "Diaphragm Centers of Mass Displacements",
            "Diaphragm Centers Of Mass Displacements",
            "Story Diaphragm Displacements",
            "Story Diaphragm Center of Mass Displacements",
            "Story/Diaphragm Displacements"
        };

        internal static readonly string[] StoryDisplacementTableNames =
        {
            "Story Displacements",
            "Story Max Over Avg Displacements"
        };

        // ─── Đọc bảng ─────────────────────────────────────────────────────────────

        public static List<Dictionary<string, string>> ReadTable(cSapModel sap, string tableKey, string outputCase)
        {
            TrySelectComboForDatabaseTables(sap, outputCase);

            string[] fieldKeyList = null;
            string groupName = "";
            int tableVersion = 0;
            string[] fieldsKeysIncluded = null;
            int numberRecords = 0;
            string[] tableData = null;

            sap.DatabaseTables.GetTableForDisplayArray(tableKey, ref fieldKeyList, groupName,
                ref tableVersion, ref fieldsKeysIncluded, ref numberRecords, ref tableData);

            var rows = new List<Dictionary<string, string>>();
            if (fieldsKeysIncluded == null || tableData == null || fieldsKeysIncluded.Length == 0)
                return rows;

            int cols = fieldsKeysIncluded.Length;
            for (int r = 0; r < numberRecords; r++)
            {
                var dict = new Dictionary<string, string>(cols, StringComparer.OrdinalIgnoreCase);
                for (int c = 0; c < cols; c++)
                {
                    int idx = r * cols + c;
                    dict[fieldsKeysIncluded[c]] = idx < tableData.Length ? tableData[idx] : "";
                }
                rows.Add(dict);
            }
            return rows;
        }

        /// <summary>
        /// Thử nhiều tên bảng cho đến khi có dữ liệu, trả về bảng đầu tiên không rỗng.
        /// </summary>
        public static List<Dictionary<string, string>> ReadTableWithFallback(
            cSapModel sap, string[] tableNames, string outputCase)
        {
            foreach (var name in tableNames)
            {
                try
                {
                    var table = ReadTable(sap, name, outputCase);
                    if (table.Count > 0) return table;
                }
                catch { /* Bảng không tồn tại trong bản ETABS này, thử tên khác */ }
            }
            return new List<Dictionary<string, string>>();
        }

        // ─── Helpers truy xuất field ──────────────────────────────────────────────

        public static string Get(Dictionary<string, string> row, params string[] keys)
        {
            // Ưu tiên 1: khớp chính xác
            foreach (var key in keys)
                if (row.TryGetValue(key, out var v)) return v;

            // Ưu tiên 2: so sánh sau khi chuẩn hóa (bỏ ký tự đặc biệt / đơn vị)
            foreach (var key in keys)
            {
                string nk = NormalizeKey(key);
                foreach (var kv in row)
                {
                    string fk = NormalizeKey(kv.Key);
                    if (fk == nk) return kv.Value;

                    if (nk == "p" && (fk == "p" || fk == "pkn" || fk == "pkip" || fk == "pnewton")) return kv.Value;
                    if (nk == "vx" && (fk == "vx" || fk == "vxkn" || fk == "vxkip")) return kv.Value;
                    if (nk == "vy" && (fk == "vy" || fk == "vykn" || fk == "vykip")) return kv.Value;
                }
            }

            // Ưu tiên 3: so khớp gần đúng (chỉ áp dụng với key >= 2 ký tự)
            foreach (var key in keys)
            {
                string nk = NormalizeKey(key);
                if (nk.Length < 2) continue;
                foreach (var kv in row)
                {
                    string fk = NormalizeKey(kv.Key);
                    if (fk.Length >= 2 && (fk.Contains(nk) || nk.Contains(fk)))
                        return kv.Value;
                }
            }

            return "";
        }

        public static double GetDouble(Dictionary<string, string> row, params string[] keys)
            => ParseDouble(Get(row, keys));

        public static double ParseDouble(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return 0.0;
            s = s.Trim().Replace(" ", "");

            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var v)) return v;
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out v)) return v;

            var s2 = s.Replace(",", "");
            if (double.TryParse(s2, NumberStyles.Any, CultureInfo.InvariantCulture, out v)) return v;

            s2 = s.Replace(".", "").Replace(",", ".");
            if (double.TryParse(s2, NumberStyles.Any, CultureInfo.InvariantCulture, out v)) return v;

            return 0.0;
        }

        public static string GetAvailableFields(IEnumerable<Dictionary<string, string>> rows)
        {
            foreach (var row in rows) return string.Join(", ", row.Keys);
            return "";
        }

        // ─── Private helpers ──────────────────────────────────────────────────────

        private static void TrySelectComboForDatabaseTables(cSapModel sap, string comboName)
        {
            if (string.IsNullOrWhiteSpace(comboName)) return;
            try
            {
                sap.Results.Setup.DeselectAllCasesAndCombosForOutput();
                sap.Results.Setup.SetComboSelectedForOutput(comboName);

                object db = sap.DatabaseTables;
                Type t = db.GetType();

                // Gọi tất cả overload phổ biến qua reflection để tương thích nhiều bản ETABSv1.dll.
                TryInvokeFlexible(db, t, "DeselectAllLoadCasesAndCombosForDisplay", comboName);
                TryInvokeFlexible(db, t, "DeselectAllCasesAndCombosForDisplay", comboName);
                TryInvokeFlexible(db, t, "SetLoadCombinationsSelectedForDisplay", comboName);
                TryInvokeFlexible(db, t, "SetLoadCombinationSelectedForDisplay", comboName);
                TryInvokeFlexible(db, t, "SetLoadCasesSelectedForDisplay", comboName);
                TryInvokeFlexible(db, t, "SetLoadCaseSelectedForDisplay", comboName);
            }
            catch { /* Không chặn nếu DLL không hỗ trợ */ }
        }

        private static void TryInvokeFlexible(object target, Type type, string methodName, string comboName)
        {
            foreach (var mi in type.GetMethods(BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance))
            {
                if (!string.Equals(mi.Name, methodName, StringComparison.OrdinalIgnoreCase)) continue;
                try
                {
                    var ps = mi.GetParameters();
                    var args = new object[ps.Length];
                    for (int i = 0; i < ps.Length; i++)
                    {
                        Type pt = ps[i].ParameterType;
                        Type bt = pt.IsByRef ? pt.GetElementType() : pt;

                        if (bt == typeof(int)) args[i] = 1;
                        else if (bt == typeof(string[])) args[i] = new[] { comboName };
                        else if (bt == typeof(string)) args[i] = comboName;
                        else if (bt == typeof(bool)) args[i] = true;
                        else args[i] = null;
                    }
                    mi.Invoke(target, args);
                    return;
                }
                catch { /* Thử overload tiếp theo */ }
            }
        }

        private static string NormalizeKey(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            var chars = new System.Text.StringBuilder();
            foreach (char ch in s.ToLowerInvariant())
                if (char.IsLetterOrDigit(ch)) chars.Append(ch);
            return chars.ToString();
        }
    }
}