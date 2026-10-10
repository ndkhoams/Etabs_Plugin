using ETABSv1;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Etabs_Ultimate_Tools
{
    public enum DisplacementSource
    {
        Diaphragm,
        Story
    }

    public enum DiaphragmAggregationMode
    {
        Max,
        Average
    }

    public static class EtabsTableReader
    {
        public static string LastDisplacementReadDiagnostic { get; private set; } = "";
        public static string LastDisplacementTableName { get; private set; } = "";
        public static string LastTableReadDiagnostic { get; private set; } = "";

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

        internal static readonly string[] DiaphragmDriftTableNames =
        {
            "Diaphragm Max Over Avg Drifts",
            "Diaphragm Max Over Average Drifts",
            "Diaphragm Drifts"
        };

        internal static readonly string[] StoryDriftTableNames =
        {
            "Story Drifts",
            "Story Drift"
        };

        internal static readonly string[] StoryDisplacementTableNames =
        {
            "Story Displacements",
            "Story Max Over Avg Displacements"
        };

        public static string[] GetDisplacementTableNames(DisplacementSource source)
            => source == DisplacementSource.Story
                ? StoryDisplacementTableNames
                : DiaphragmDisplacementTableNames;

        public static Dictionary<string, double> ReadStoryDisplacements(
            cSapModel sap, string combo, string dir, DisplacementSource source,
            DiaphragmAggregationMode aggregationMode = DiaphragmAggregationMode.Max)
        {
            return ReadStoryDisplacementInfo(sap, combo, dir, source, aggregationMode).Item1;
        }

        public static Dictionary<string, double> ReadDiaphragmDrifts(
            cSapModel sap, string combo, string dir, DiaphragmAggregationMode aggregationMode)
        {
            var table = ReadTableWithFallback(sap, DiaphragmDriftTableNames, combo);
            var storyValues = new Dictionary<string, Dictionary<string, Tuple<double, double>>>(
                StringComparer.OrdinalIgnoreCase);

            foreach (var row in table)
            {
                string story = Get(row, "Story", "StoryName", "Story Name", "Level").Trim();
                if (string.IsNullOrWhiteSpace(story)) continue;

                string outputCase = Get(row,
                    "Output Case", "OutputCase", "Load Case", "LoadCase", "Case", "Combo", "Combination");
                if (!EtabsHelper.IsSameOrEnvelopeCase(outputCase, combo)) continue;

                string item = Get(row, "Item", "Diaphragm", "Diaphragm Name", "Name").Trim();
                string rowDirection = Get(row, "Direction", "Dir").Trim();
                if (string.IsNullOrWhiteSpace(rowDirection))
                    rowDirection = item.EndsWith(" Y", StringComparison.OrdinalIgnoreCase) ? "Y" :
                        item.EndsWith(" X", StringComparison.OrdinalIgnoreCase) ? "X" : "";
                if (!rowDirection.StartsWith(dir, StringComparison.OrdinalIgnoreCase)) continue;

                double maxDrift = Math.Abs(GetDouble(row, "Max Drift", "Maximum Drift"));
                double averageDrift = Math.Abs(GetDouble(row, "Avg Drift", "Average Drift"));
                if (maxDrift <= 0 && averageDrift <= 0) continue;

                string diaphragm = Get(row,
                    "Diaphragm", "Diaphragm ID", "Diaphragm Name", "Name").Trim();
                if (string.IsNullOrWhiteSpace(diaphragm))
                    diaphragm = item.Length > 2 ? item.Substring(0, item.Length - 2).Trim() : item;
                if (string.IsNullOrWhiteSpace(diaphragm)) diaphragm = story;

                if (!storyValues.TryGetValue(story, out var diaphragmValues))
                {
                    diaphragmValues = new Dictionary<string, Tuple<double, double>>(
                        StringComparer.OrdinalIgnoreCase);
                    storyValues[story] = diaphragmValues;
                }

                if (diaphragmValues.TryGetValue(diaphragm, out var current))
                    diaphragmValues[diaphragm] = Tuple.Create(
                        Math.Max(current.Item1, maxDrift), Math.Max(current.Item2, averageDrift));
                else
                    diaphragmValues[diaphragm] = Tuple.Create(maxDrift, averageDrift);
            }

            var result = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            foreach (var story in storyValues)
            {
                var values = story.Value.Values;
                result[story.Key] = aggregationMode == DiaphragmAggregationMode.Average
                    ? values.Where(value => value.Item2 > 0).Select(value => value.Item2).DefaultIfEmpty(0).Average()
                    : values.Max(value => value.Item1);
            }

            if (result.Count > 0)
            {
                LastDisplacementTableName = "Diaphragm Max Over Avg Drifts";
                LastDisplacementReadDiagnostic = "Nguồn=Diaphragm; bảng=" + LastDisplacementTableName +
                    "; tổ hợp=" + combo + "; phương=" + dir + "; phương pháp=" + aggregationMode;
            }

            return result;
        }

        public static Tuple<Dictionary<string, double>, Dictionary<string, string>> ReadStoryDisplacementInfo(
            cSapModel sap, string combo, string dir, DisplacementSource source,
            DiaphragmAggregationMode aggregationMode = DiaphragmAggregationMode.Max)
        {
            if (sap.SetPresentUnits(eUnits.kN_m_C) != 0)
                throw new InvalidOperationException("Không thể đặt đơn vị ETABS về kN-m trước khi đọc chuyển vị.");

            var bestWithCase = new Dictionary<string, List<double>>(StringComparer.OrdinalIgnoreCase);
            var bestWithCaseName = new Dictionary<string, List<Tuple<double, string>>>(StringComparer.OrdinalIgnoreCase);
            var observedOutputCases = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var observedStepTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var tableNames = new List<string>(GetDisplacementTableNames(source));
            var readErrors = new List<string>();
            int totalRows = 0;
            string lastTable = "", lastFields = "";

            foreach (string tableKey in GetAvailableDisplacementTableKeys(sap, source))
                if (!tableNames.Exists(name =>
                    string.Equals(name, tableKey, StringComparison.OrdinalIgnoreCase)))
                    tableNames.Add(tableKey);

            foreach (var tableName in tableNames)
            {
                List<Dictionary<string, string>> table;
                try { table = ReadTable(sap, tableName, combo); }
                catch (Exception ex)
                {
                    readErrors.Add(tableName + ": " + ex.Message);
                    continue;
                }
                totalRows += table.Count;
                if (table.Count > 0)
                {
                    lastTable = tableName;
                    lastFields = GetAvailableFields(table);
                }

                foreach (var row in table)
                {
                    string story = Get(row, "Story", "StoryName", "Story Name", "Level");
                    if (string.IsNullOrWhiteSpace(story)) continue;
                    story = story.Trim();

                    string outputCase = Get(row,
                        "Output Case", "OutputCase", "Load Case", "LoadCase", "Case", "Combo", "Combination");
                    string stepType = Get(row, "Step Type", "StepType", "Step");
                    if (!string.IsNullOrWhiteSpace(outputCase)) observedOutputCases.Add(outputCase.Trim());
                    if (!string.IsNullOrWhiteSpace(stepType)) observedStepTypes.Add(stepType.Trim());
                    double displacement = ReadDirectionalDisplacement(row, dir, source);

                    string diaphragmName = Get(row,
                        "Diaphragm", "Diaphragm ID", "Diaphragm Name", "DiaphragmName",
                        "Name", "Label", "Object", "Story Diaphragm");
                    if (!string.IsNullOrWhiteSpace(diaphragmName)) diaphragmName = diaphragmName.Trim();

                    if (EtabsHelper.IsSameOrEnvelopeCase(outputCase, combo))
                    {
                        AddStoryValue(bestWithCase, story, displacement);
                        if (!string.IsNullOrWhiteSpace(diaphragmName))
                            AddStoryName(bestWithCaseName, story, diaphragmName, displacement);
                    }
                }

                if (bestWithCase.Values.Any(values => values.Any(v => Math.Abs(v) > 1e-12)))
                {
                    LastDisplacementTableName = tableName;
                    LastDisplacementReadDiagnostic =
                        "Nguồn=" + source + "; bảng=" + tableName + "; số dòng=" + totalRows +
                        "; trường=" + lastFields;
                    var named = AggregateStoryNames(bestWithCaseName);
                    return Tuple.Create(AggregateStoryValues(bestWithCase, aggregationMode), named);
                }

            }
                        var result = AggregateStoryValues(bestWithCase, aggregationMode);
                        var resultName = AggregateStoryNames(bestWithCaseName);
            bool hasNonZero = result.Values.Any(value => Math.Abs(value) > 1e-12);
            LastDisplacementTableName = lastTable;
                        LastDisplacementReadDiagnostic = bestWithCase.Count == 0
                                ? "Không đọc được chuyển vị cho tổ hợp " + combo + "; OutputCase quan sát=" +
                                    string.Join(", ", observedOutputCases.Take(10)) + "; StepType=" +
                                    string.Join(", ", observedStepTypes.Take(10)) +
                                    "; nguồn=" + source + "; bảng cuối=" + lastTable + "; số dòng=" + totalRows + "; trường=" + lastFields +
                                    (readErrors.Count > 0 ? "; lỗi đọc bảng=" + string.Join(" | ", readErrors) : "")
                                : "Nguồn=" + source + "; bảng cuối=" + lastTable + "; số dòng=" + totalRows +
                                    "; trường=" + lastFields + "; có giá trị khác 0=" + hasNonZero;
            return Tuple.Create(result, resultName);
        }

        private static List<string> GetAvailableDisplacementTableKeys(
            cSapModel sap, DisplacementSource source)
        {
            int count = 0;
            string[] keys = null, names = null;
            int[] importTypes = null;
            var matches = new List<string>();

            try
            {
                if (sap.DatabaseTables.GetAvailableTables(
                    ref count, ref keys, ref names, ref importTypes) != 0 || keys == null)
                    return matches;

                for (int i = 0; i < count && i < keys.Length; i++)
                {
                    string descriptor = keys[i] + " " +
                        (names != null && i < names.Length ? names[i] : "");
                    bool isDisplacementTable = descriptor.IndexOf(
                        "Displacement", StringComparison.OrdinalIgnoreCase) >= 0;
                    bool isSourceTable = source == DisplacementSource.Story
                        ? descriptor.IndexOf("Story", StringComparison.OrdinalIgnoreCase) >= 0
                        : descriptor.IndexOf("Diaphragm", StringComparison.OrdinalIgnoreCase) >= 0;

                    if (isDisplacementTable && isSourceTable && !string.IsNullOrWhiteSpace(keys[i]))
                        matches.Add(keys[i]);
                }

            }
            catch
            {
                return matches;
            }

            return matches;
        }

        private static void AddStoryValue(
            Dictionary<string, List<double>> values, string key, double value)
        {
            if (!values.TryGetValue(key, out var list))
            {
                list = new List<double>();
                values[key] = list;
            }

            list.Add(value);
        }

        private static void AddStoryName(
            Dictionary<string, List<Tuple<double, string>>> names, string key, string name, double value)
        {
            double magnitude = Math.Abs(value);
            if (!names.TryGetValue(key, out var list))
            {
                list = new List<Tuple<double, string>>();
                names[key] = list;
            }

            list.Add(Tuple.Create(magnitude, name));
        }

        private static Dictionary<string, double> AggregateStoryValues(
            Dictionary<string, List<double>> values, DiaphragmAggregationMode mode)
        {
            var result = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            foreach (var pair in values)
            {
                if (pair.Value == null || pair.Value.Count == 0) continue;
                var magnitudes = pair.Value.Select(v => Math.Abs(v)).ToList();
                result[pair.Key] = mode == DiaphragmAggregationMode.Average
                    ? magnitudes.Average()
                    : magnitudes.Max();
            }

            return result;
        }

        private static Dictionary<string, string> AggregateStoryNames(
            Dictionary<string, List<Tuple<double, string>>> names)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var pair in names)
            {
                if (pair.Value == null || pair.Value.Count == 0) continue;
                var selected = pair.Value.OrderByDescending(item => item.Item1).FirstOrDefault();
                if (selected != null)
                    result[pair.Key] = selected.Item2;
            }

            return result;
        }

        private static double ReadDirectionalDisplacement(
            Dictionary<string, string> row, string dir, DisplacementSource source)
        {
            double displacement = dir.Equals("X", StringComparison.OrdinalIgnoreCase)
                ? GetDouble(row,
                    "UX", "Ux", "UX m", "Ux m", "UX mm", "Ux mm",
                    "U1", "U1 m", "U1 mm", "X", "X m", "X mm",
                    "X-Displ", "X Displ", "Displ X", "Translation X",
                    "Global X", "GlobalX", "X Translation")
                : GetDouble(row,
                    "UY", "Uy", "UY m", "Uy m", "UY mm", "Uy mm",
                    "U2", "U2 m", "U2 mm", "Y", "Y m", "Y mm",
                    "Y-Displ", "Y Displ", "Displ Y", "Translation Y",
                    "Global Y", "GlobalY", "Y Translation");

            if (Math.Abs(displacement) > 1e-12 || source != DisplacementSource.Story)
                return displacement;

            string rowDirection = Get(row, "Direction", "Dir");
            if (string.IsNullOrWhiteSpace(rowDirection) ||
                !rowDirection.Trim().StartsWith(dir, StringComparison.OrdinalIgnoreCase))
                return 0.0;

            return GetDouble(row, "Maximum Displacement", "Maximum", "Max Displacement", "Max");
        }

        public static List<Dictionary<string, string>> ReadTable(cSapModel sap, string tableKey, string outputCase)
        {
            TrySelectComboForDatabaseTables(sap, outputCase);

            string[] fieldKeyList = null;
            string groupName = "";
            int tableVersion = 0;
            string[] fieldsKeysIncluded = null;
            int numberRecords = 0;
            string[] tableData = null;

            int ret = sap.DatabaseTables.GetTableForDisplayArray(tableKey, ref fieldKeyList, groupName,
                ref tableVersion, ref fieldsKeysIncluded, ref numberRecords, ref tableData);
            if (ret != 0)
                throw new InvalidOperationException("ETABS không đọc được bảng '" + tableKey +
                    "' cho tổ hợp '" + outputCase + "' (return code " + ret + ").");

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

        /// <summary>Thử nhiều tên bảng cho đến khi có dữ liệu.</summary>
        public static List<Dictionary<string, string>> ReadTableWithFallback(
            cSapModel sap, string[] tableNames, string outputCase)
        {
            var errors = new List<string>();
            foreach (var name in tableNames)
            {
                try
                {
                    var table = ReadTable(sap, name, outputCase);
                    if (table.Count > 0)
                    {
                        LastTableReadDiagnostic = errors.Count == 0
                            ? "Đọc bảng '" + name + "' cho tổ hợp '" + outputCase + "'."
                            : "Đã dùng bảng '" + name + "' cho tổ hợp '" + outputCase +
                              "'; alias trước đó lỗi: " + string.Join(" | ", errors);
                        return table;
                    }
                }
                catch (Exception ex)
                {
                    errors.Add("'" + name + "': " + ex.Message);
                }
            }
            LastTableReadDiagnostic = "Không có dữ liệu từ các bảng [" + string.Join(", ", tableNames) +
                "] cho tổ hợp '" + outputCase + "'." +
                (errors.Count > 0 ? " Lỗi: " + string.Join(" | ", errors) : "");
            return new List<Dictionary<string, string>>();
        }

        public static string Get(Dictionary<string, string> row, params string[] keys)
        {
            foreach (var key in keys)
                if (row.TryGetValue(key, out var v)) return v;

            var normMap = new Dictionary<string, string>(row.Count);
            foreach (var kv in row)
            {
                string fk = NormalizeKey(kv.Key);
                if (fk.Length > 0 && !normMap.ContainsKey(fk)) normMap[fk] = kv.Value;
            }

            foreach (var key in keys)
            {
                string nk = NormalizeKey(key);
                if (normMap.TryGetValue(nk, out var v)) return v;

                if (nk == "p" && (normMap.TryGetValue("pkn", out v) ||
                                  normMap.TryGetValue("pkip", out v) ||
                                  normMap.TryGetValue("pnewton", out v))) return v;
                if (nk == "vx" && (normMap.TryGetValue("vxkn", out v) ||
                                   normMap.TryGetValue("vxkip", out v))) return v;
                if (nk == "vy" && (normMap.TryGetValue("vykn", out v) ||
                                   normMap.TryGetValue("vykip", out v))) return v;
            }

            foreach (var key in keys)
            {
                string nk = NormalizeKey(key);
                if (nk.Length < 2) continue;
                foreach (var kv in normMap)
                {
                    if (kv.Key.Length >= 2 && (kv.Key.Contains(nk) || nk.Contains(kv.Key)))
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

        private static void TrySelectComboForDatabaseTables(cSapModel sap, string comboName)
        {
            if (string.IsNullOrWhiteSpace(comboName)) return;
            sap.Results.Setup.DeselectAllCasesAndCombosForOutput();
            int outputRet = sap.Results.Setup.SetComboSelectedForOutput(comboName);
            if (outputRet != 0)
                throw new InvalidOperationException("Không chọn được tổ hợp '" + comboName +
                    "' cho kết quả ETABS (return code " + outputRet + ").");

            string[] noCases = new string[0];
            int clearCasesRet = sap.DatabaseTables.SetLoadCasesSelectedForDisplay(ref noCases);
            if (clearCasesRet != 0)
                throw new InvalidOperationException("Không xóa được load case cũ khỏi Database Tables " +
                    "(return code " + clearCasesRet + ").");

            string[] selectedCombos = { comboName };
            int comboRet = sap.DatabaseTables.SetLoadCombinationsSelectedForDisplay(ref selectedCombos);
            if (comboRet != 0)
                throw new InvalidOperationException("Không chọn được tổ hợp '" + comboName +
                    "' trong Database Tables (return code " + comboRet + ").");
        }

        private static string NormalizeKey(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            var chars = new System.Text.StringBuilder(s.Length);
            foreach (char ch in s.ToLowerInvariant())
                if (char.IsLetterOrDigit(ch)) chars.Append(ch);
            return chars.ToString();
        }
    }
}
