using ETABSv1;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;

namespace CheckModelPlugin
{
    public static class EtabsTableReader
    {
        public static List<Dictionary<string, string>> ReadTable(cSapModel sap, string tableKey, string outputCase)
        {
            // Với DatabaseTables, cần chọn combo cho Display Tables, không chỉ Results.Setup.
            // Dùng reflection để tương thích nhiều bản ETABSv1.dll.
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
            if (fieldsKeysIncluded == null || tableData == null || fieldsKeysIncluded.Length == 0) return rows;

            int cols = fieldsKeysIncluded.Length;
            for (int r = 0; r < numberRecords; r++)
            {
                var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                for (int c = 0; c < cols; c++)
                {
                    int idx = r * cols + c;
                    dict[fieldsKeysIncluded[c]] = idx < tableData.Length ? tableData[idx] : "";
                }
                rows.Add(dict);
            }
            return rows;
        }

        private static void TrySelectComboForDatabaseTables(cSapModel sap, string comboName)
        {
            if (string.IsNullOrWhiteSpace(comboName)) return;

            try
            {
                // Results.Setup giúp một số bản ETABS đổi combo cho Display Tables.
                // DatabaseTables lại có bộ chọn riêng cho Display; vì vậy gọi cả hai.
                sap.Results.Setup.DeselectAllCasesAndCombosForOutput();
                sap.Results.Setup.SetComboSelectedForOutput(comboName);

                object db = sap.DatabaseTables;
                Type t = db.GetType();

                // ETABS/SAP/SAFE các đời DLL khác nhau có chữ ký hàm khác nhau:
                // - DeselectAllLoadCasesAndCombosForDisplay()
                // - SetLoadCombinationsSelectedForDisplay(ref int NumberItems, ref string[] MyName)
                // - SetLoadCombinationsSelectedForDisplay(int NumberItems, string[] MyName)
                // - SetLoadCombinationSelectedForDisplay(string Name)
                // Dùng reflection mềm để bắt được mọi overload phổ biến.
                TryInvokeFlexible(db, t, "DeselectAllLoadCasesAndCombosForDisplay", comboName);
                TryInvokeFlexible(db, t, "DeselectAllCasesAndCombosForDisplay", comboName);
                TryInvokeFlexible(db, t, "SetLoadCombinationsSelectedForDisplay", comboName);
                TryInvokeFlexible(db, t, "SetLoadCombinationSelectedForDisplay", comboName);
                TryInvokeFlexible(db, t, "SetLoadCasesSelectedForDisplay", comboName);
                TryInvokeFlexible(db, t, "SetLoadCaseSelectedForDisplay", comboName);
            }
            catch
            {
                // Không chặn xuất nếu DLL không hỗ trợ các hàm display selection.
            }
        }

        private static void TryInvokeFlexible(object target, Type type, string methodName, string comboName)
        {
            foreach (var mi in type.GetMethods(BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance))
            {
                if (!string.Equals(mi.Name, methodName, StringComparison.OrdinalIgnoreCase)) continue;

                try
                {
                    var ps = mi.GetParameters();
                    object[] args = new object[ps.Length];

                    for (int i = 0; i < ps.Length; i++)
                    {
                        Type pt = ps[i].ParameterType;
                        Type baseType = pt.IsByRef ? pt.GetElementType() : pt;

                        if (baseType == typeof(int))
                            args[i] = 1;
                        else if (baseType == typeof(string[]))
                            args[i] = new string[] { comboName };
                        else if (baseType == typeof(string))
                            args[i] = comboName;
                        else if (baseType == typeof(bool))
                            args[i] = true;
                        else
                            args[i] = null;
                    }

                    mi.Invoke(target, args);
                    return;
                }
                catch
                {
                    // Thử overload khác.
                }
            }
        }

        public static string Get(Dictionary<string, string> row, params string[] keys)
        {
            // Ưu tiên 1: đúng tên field ETABS trả về, ví dụ: Story, Output Case, Location, P, VX, VY.
            foreach (var key in keys)
            {
                if (row.TryGetValue(key, out var v)) return v;
            }

            // Ưu tiên 2: so khớp sau khi bỏ ký tự đặc biệt / xuống dòng đơn vị.
            // Ví dụ field API có thể là "P\nkN", "P kN", "VX kN".
            foreach (var key in keys)
            {
                string nk = NormalizeKey(key);
                foreach (var kv in row)
                {
                    string fk = NormalizeKey(kv.Key);
                    if (fk == nk) return kv.Value;

                    // Riêng các field nội lực ngắn P, VX, VY cho phép có đơn vị phía sau.
                    // Không cho phép match bừa với OutputCase/StepType.
                    if (nk == "p" && (fk == "p" || fk == "pkn" || fk == "pkip" || fk == "pnewton")) return kv.Value;
                    if (nk == "vx" && (fk == "vx" || fk == "vxkn" || fk == "vxkip")) return kv.Value;
                    if (nk == "vy" && (fk == "vy" || fk == "vykn" || fk == "vykip")) return kv.Value;
                }
            }

            // Ưu tiên 3: so khớp gần đúng, nhưng KHÔNG áp dụng với key quá ngắn như "P".
            foreach (var key in keys)
            {
                string nk = NormalizeKey(key);
                if (nk.Length < 2) continue;

                foreach (var kv in row)
                {
                    string fk = NormalizeKey(kv.Key);
                    if (fk.Length < 2) continue;
                    if (fk.Contains(nk) || nk.Contains(fk)) return kv.Value;
                }
            }

            return "";
        }

        public static string GetAvailableFields(IEnumerable<Dictionary<string, string>> rows)
        {
            foreach (var row in rows) return string.Join(", ", row.Keys);
            return "";
        }

        public static double GetDouble(Dictionary<string, string> row, params string[] keys)
        {
            var s = Get(row, keys);
            return ParseDouble(s);
        }

        public static double ParseDouble(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return 0.0;
            s = s.Trim().Replace(" ", "");

            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var v)) return v;
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out v)) return v;

            // Trường hợp máy dùng dấu phẩy/dấu chấm khác nhau.
            var s2 = s.Replace(",", "");
            if (double.TryParse(s2, NumberStyles.Any, CultureInfo.InvariantCulture, out v)) return v;
            s2 = s.Replace(".", "").Replace(",", ".");
            if (double.TryParse(s2, NumberStyles.Any, CultureInfo.InvariantCulture, out v)) return v;
            return 0.0;
        }

        private static string NormalizeKey(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            var chars = new List<char>();
            foreach (char ch in s.ToLowerInvariant())
            {
                if (char.IsLetterOrDigit(ch)) chars.Add(ch);
            }
            return new string(chars.ToArray());
        }
    }
}
