using ETABSv1;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Etabs_Ultimate_Tools
{
    /// <summary>
    /// Đọc / đổi tên design strip thông qua Interactive Database Tables của ETABS.
    ///
    /// Lưu ý quan trọng: ETABS OAPI KHÔNG cho phép lấy "strip đang chọn" trực tiếp
    /// (SelectObj.GetSelected chỉ trả về Point/Frame/Cable/Tendon/Area/Solid/Link, không có
    /// design strip; cGroup.GetAssignments cũng dùng bộ loại đó). Vì vậy ta lọc theo Group
    /// ngay trong bảng database (tham số GroupName).
    ///
    /// Toạ độ strip: bảng rên tên (ví dụ overwrites) thường không có toạ độ, nên
    /// ta tự quét các bảng strip khác (đọc bằng GetTableForDisplayArray – chạy được cả với
    /// bảng chỉ-xem) để tìm bảng có toạ độ điểm, rồi tính tâm mỗi strip.
    /// </summary>
    internal class StripRenamer
    {
        public enum SortMode
        {
            RowLtoR_TtoB = 0,   // Hàng ngang: Trái→Phải, Trên→Dưới
            RowRtoL_BtoT = 1,   // Hàng ngang ngược: Phải→Trái, Dưới→Trên
            ColTtoB_LtoR = 2,   // Cột dọc: Trên→Dưới, Trái→Phải
            ColBtoT_RtoL = 3    // Cột dọc ngược: Dưới→Trên, Phải→Trái
        }

        public class StripItem
        {
            public string OldName;
            public string NewName;
            public double X;
            public double Y;
            public int RecordCount;
        }

        public class TableInfo
        {
            public string Key;
            public string Name;
            public int ImportType;
        }

        public string TableKey = "";
        public string GroupName = "";
        public int TableVersion;
        public int NumberFields;
        public string[] FieldKeys = new string[0];
        public int NumberRecords;
        public string[] TableData = new string[0];

        public int NameCol = -1;
        public int XCol = -1;
        public int YCol = -1;

        // Toạ độ lấy từ bảng khác (name -> {x, y}).
        public Dictionary<string, double[]> CoordMap = new Dictionary<string, double[]>(StringComparer.Ordinal);
        public string CoordSource = "";

        /// <summary>Liệt kê toàn bộ bảng trong database của model (kèm importType).</summary>
        public static List<TableInfo> FindAllTables(cSapModel sap)
        {
            var list = new List<TableInfo>();
            int n = 0;
            string[] keys = null, names = null;
            int[] importType = null;
            bool[] isEmpty = null;
            sap.DatabaseTables.GetAllTables(ref n, ref keys, ref names, ref importType, ref isEmpty);
            if (keys == null) return list;
            for (int i = 0; i < keys.Length; i++)
            {
                list.Add(new TableInfo
                {
                    Key = keys[i] ?? "",
                    Name = (names != null && i < names.Length) ? (names[i] ?? "") : "",
                    ImportType = (importType != null && i < importType.Length) ? importType[i] : -1
                });
            }
            return list;
        }

        /// <summary>Các bảng có chữ "strip" trong key/tên (để gợi ý cho người dùng).</summary>
        public static List<TableInfo> FindStripTables(cSapModel sap)
        {
            var list = new List<TableInfo>();
            foreach (var t in FindAllTables(sap))
            {
                if (t.Key.IndexOf("strip", StringComparison.OrdinalIgnoreCase) >= 0
                    || t.Name.IndexOf("strip", StringComparison.OrdinalIgnoreCase) >= 0)
                    list.Add(t);
            }
            return list;
        }

        /// <summary>Đọc bảng để chỉnh sửa, lọc theo group.</summary>
        public void ReadGroup(cSapModel sap, string tableKey, string groupName)
        {
            TableKey = tableKey;
            GroupName = groupName ?? "";
            int tableVersion = 0;
            string[] fieldKeys = null;
            int numRecords = 0;
            string[] data = null;

            int ret = sap.DatabaseTables.GetTableForEditingArray(
                tableKey, GroupName, ref tableVersion, ref fieldKeys, ref numRecords, ref data);
            if (ret != 0)
                throw new Exception("Không đọc được bảng \"" + tableKey + "\" (mã lỗi " + ret + ").\n" +
                    "Có thể key tên bảng chưa đúng hoặc bảng này không cho phép chỉnh sửa (importType = 0).");

            TableVersion = tableVersion;
            FieldKeys = fieldKeys ?? new string[0];
            NumberFields = FieldKeys.Length;
            NumberRecords = numRecords;
            TableData = data ?? new string[0];

            NameCol = FindNameCol(FieldKeys, false);
            XCol = FindCoordCol(FieldKeys, "x");
            YCol = FindCoordCol(FieldKeys, "y");
        }

        /// <summary>Quét các bảng strip khác để tìm toạ độ điểm, tính tâm mỗi strip. Trả về true nếu có.</summary>
        public bool LoadCoordinatesFromStripTables(cSapModel sap, string groupName, string excludeTableKey)
        {
            CoordMap = new Dictionary<string, double[]>(StringComparer.Ordinal);
            CoordSource = "";
            var candidates = FindStripTables(sap)
                .OrderByDescending(t => ScoreCoordTable(t)).ToList();
            foreach (var t in candidates)
            {
                if (!string.IsNullOrEmpty(excludeTableKey)
                    && string.Equals(t.Key, excludeTableKey, StringComparison.OrdinalIgnoreCase)) continue;
                try
                {
                    if (TryLoadCoordFromTable(sap, t.Key, groupName)) { CoordSource = t.Key; return true; }
                }
                catch { /* bỏ qua bảng không đọc được */ }
            }
            return false;
        }

        private static int ScoreCoordTable(TableInfo t)
        {
            string s = ((t.Key ?? "") + " " + (t.Name ?? "")).ToLowerInvariant();
            int score = 0;
            if (s.Contains("connectivity")) score += 50;
            if (s.Contains("geometry")) score += 45;
            if (s.Contains("object")) score += 20;
            if (s.Contains("point")) score += 15;
            if (s.Contains("general")) score += 10;
            if (s.Contains("overwrite")) score -= 40;
            if (s.Contains("reinforc") || s.Contains("result") || s.Contains("forces")) score -= 30;
            return score;
        }

        private bool TryLoadCoordFromTable(cSapModel sap, string tableKey, string groupName)
        {
            int tableVersion = 0;
            string[] fieldKeyList = new string[0];
            string[] fieldsIncluded = null;
            int numRecords = 0;
            string[] data = null;

            int ret = sap.DatabaseTables.GetTableForDisplayArray(
                tableKey, ref fieldKeyList, groupName ?? "", ref tableVersion,
                ref fieldsIncluded, ref numRecords, ref data);
            if (ret != 0 || fieldsIncluded == null || data == null) return false;

            int nf = fieldsIncluded.Length;
            if (nf <= 0) return false;
            int nameC = FindNameCol(fieldsIncluded, true);
            int xC = FindCoordCol(fieldsIncluded, "x");
            int yC = FindCoordCol(fieldsIncluded, "y");
            if (nameC < 0 || xC < 0 || yC < 0) return false;

            var acc = new Dictionary<string, double[]>(StringComparer.Ordinal); // name -> {sumX, sumY, count}
            int rows = data.Length / nf;
            for (int r = 0; r < rows; r++)
            {
                string nm = GetCell(data, r, nf, nameC);
                if (string.IsNullOrWhiteSpace(nm)) continue;
                double x, y;
                bool hx = TryNum(GetCell(data, r, nf, xC), out x);
                bool hy = TryNum(GetCell(data, r, nf, yC), out y);
                if (!hx && !hy) continue;
                double[] a;
                if (!acc.TryGetValue(nm, out a)) { a = new double[3]; acc[nm] = a; }
                if (hx) a[0] += x;
                if (hy) a[1] += y;
                a[2] += 1;
            }
            if (acc.Count == 0) return false;
            foreach (var kv in acc)
            {
                double c = kv.Value[2] > 0 ? kv.Value[2] : 1;
                CoordMap[kv.Key] = new double[] { kv.Value[0] / c, kv.Value[1] / c };
            }
            return CoordMap.Count > 0;
        }

        private static string GetCell(string[] data, int r, int nf, int c)
        {
            int idx = r * nf + c;
            if (idx < 0 || idx >= data.Length) return "";
            return data[idx] ?? "";
        }

        private static int FindNameCol(string[] fields, bool requireMatch)
        {
            int best = -1, bestScore = 0;
            for (int i = 0; i < fields.Length; i++)
            {
                string fl = (fields[i] ?? "").ToLowerInvariant().Replace(" ", "");
                int score = 0;
                if (fl == "uniquename") score = 100;
                else if (fl.Contains("strip") && fl.Contains("name")) score = 90;
                else if (fl == "strip") score = 80;
                else if (fl == "name") score = 70;
                else if (fl.Contains("name") && !fl.Contains("point") && !fl.Contains("group")
                         && !fl.Contains("story") && !fl.Contains("layer") && !fl.Contains("section")) score = 40;
                if (score > bestScore) { bestScore = score; best = i; }
            }
            if (best >= 0 && bestScore > 0) return best;
            if (requireMatch) return -1;
            return fields.Length > 0 ? 0 : -1;
        }

        private static int FindCoordCol(string[] fields, string axis)
        {
            int best = -1, bestScore = 0;
            for (int i = 0; i < fields.Length; i++)
            {
                string fl = (fields[i] ?? "").ToLowerInvariant().Replace(" ", "");
                int score = 0;
                if (fl == "global" + axis) score = 100;
                else if (fl.Contains("global") && fl.Contains(axis)) score = 90;
                else if (fl == axis) score = 80;
                else if (fl.Contains(axis + "coord") || fl.Contains("coord" + axis)) score = 70;
                else if (fl.StartsWith(axis) && fl.Length <= 3) score = 50;
                if (score > bestScore) { bestScore = score; best = i; }
            }
            return bestScore > 0 ? best : -1;
        }

        private string Cell(int r, int c)
        {
            if (NumberFields <= 0) return "";
            int idx = r * NumberFields + c;
            if (idx < 0 || idx >= TableData.Length) return "";
            return TableData[idx] ?? "";
        }

        private static bool TryNum(string s, out double v)
        {
            v = 0;
            if (string.IsNullOrWhiteSpace(s)) return false;
            s = s.Trim();
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out v)) return true;
            return double.TryParse(s.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out v);
        }

        /// <summary>Gom record theo tên strip, sắp xếp theo vị trí, gán tên mới.</summary>
        public List<StripItem> BuildPlan(int nameCol, int xCol, int yCol,
            string prefix, int startNo, int pad, SortMode mode, double tol)
        {
            if (nameCol < 0) throw new Exception("Chưa chọn cột Tên strip.");
            if (pad < 1) pad = 1;

            var map = new Dictionary<string, StripItem>(StringComparer.Ordinal);
            var order = new List<string>();
            for (int r = 0; r < NumberRecords; r++)
            {
                string oldName = Cell(r, nameCol);
                if (string.IsNullOrWhiteSpace(oldName)) continue;
                StripItem it;
                if (!map.TryGetValue(oldName, out it))
                {
                    it = new StripItem { OldName = oldName, X = 0, Y = 0, RecordCount = 0 };
                    map[oldName] = it;
                    order.Add(oldName);
                }
                double x, y;
                if (xCol >= 0 && TryNum(Cell(r, xCol), out x)) it.X += x;
                if (yCol >= 0 && TryNum(Cell(r, yCol), out y)) it.Y += y;
                it.RecordCount++;
            }

            var items = new List<StripItem>();
            foreach (var k in order)
            {
                var it = map[k];
                if ((xCol >= 0 || yCol >= 0) && it.RecordCount > 0) { it.X /= it.RecordCount; it.Y /= it.RecordCount; }
                items.Add(it);
            }

            bool hasCoord = xCol >= 0 && yCol >= 0;
            if (!hasCoord && CoordMap != null && CoordMap.Count > 0)
            {
                int matched = 0;
                foreach (var it in items)
                {
                    double[] c;
                    if (CoordMap.TryGetValue(it.OldName, out c)) { it.X = c[0]; it.Y = c[1]; matched++; }
                }
                if (matched > 0) hasCoord = true;
            }

            var sorted = SortItems(items, mode, tol, hasCoord);

            int idx = startNo;
            foreach (var it in sorted)
            {
                it.NewName = (prefix ?? "") + idx.ToString(CultureInfo.InvariantCulture).PadLeft(pad, '0');
                idx++;
            }
            return sorted;
        }

        private static List<StripItem> SortItems(List<StripItem> items, SortMode mode, double tol, bool hasCoord)
        {
            if (!hasCoord)
                return items.OrderBy(i => i.OldName, StringComparer.OrdinalIgnoreCase).ToList();
            if (tol <= 0) tol = 0.0001;
            double t = tol;
            Func<double, long> bin = v => (long)Math.Round(v / t);
            switch (mode)
            {
                case SortMode.RowLtoR_TtoB:
                    return items.OrderByDescending(i => bin(i.Y)).ThenBy(i => i.X).ToList();
                case SortMode.RowRtoL_BtoT:
                    return items.OrderBy(i => bin(i.Y)).ThenByDescending(i => i.X).ToList();
                case SortMode.ColTtoB_LtoR:
                    return items.OrderBy(i => bin(i.X)).ThenByDescending(i => i.Y).ToList();
                case SortMode.ColBtoT_RtoL:
                    return items.OrderByDescending(i => bin(i.X)).ThenBy(i => i.Y).ToList();
                default:
                    return items;
            }
        }

        /// <summary>Ghi tên mới vào model. Trả về log import.</summary>
        public string Apply(cSapModel sap, List<StripItem> plan, int nameCol)
        {
            if (nameCol < 0) throw new Exception("Chưa chọn cột Tên strip.");
            var rename = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var it in plan)
                if (!string.IsNullOrWhiteSpace(it.NewName)) rename[it.OldName] = it.NewName;

            // kiểm tra trùng tên mới
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var kv in rename)
                if (!seen.Add(kv.Value))
                    throw new Exception("Tên mới bị trùng: " + kv.Value + ". Đổi prefix hoặc số bắt đầu.");

            string[] data = (string[])TableData.Clone();
            for (int r = 0; r < NumberRecords; r++)
            {
                int idx = r * NumberFields + nameCol;
                if (idx < 0 || idx >= data.Length) continue;
                string oldName = data[idx];
                string nn;
                if (oldName != null && rename.TryGetValue(oldName, out nn))
                    data[idx] = nn;
            }

            sap.SetModelIsLocked(false);

            int tableVersion = TableVersion;
            string[] fieldKeys = FieldKeys;
            int numRecords = NumberRecords;
            int ret = sap.DatabaseTables.SetTableForEditingArray(
                TableKey, ref tableVersion, ref fieldKeys, numRecords, ref data);
            if (ret != 0) throw new Exception("SetTableForEditingArray lỗi (mã " + ret + ").");

            int fatal = 0, err = 0, warn = 0, info = 0;
            string log = "";
            int ret2 = sap.DatabaseTables.ApplyEditedTables(true, ref fatal, ref err, ref warn, ref info, ref log);
            if (ret2 != 0) throw new Exception("ApplyEditedTables lỗi (mã " + ret2 + ").");

            string summary = "Fatal=" + fatal + ", Error=" + err + ", Warn=" + warn + ", Info=" + info;
            if (!string.IsNullOrEmpty(log)) summary += "\n" + log;
            return summary;
        }
    }

    public class StripPreviewRow
    {
        public int Stt { get; set; }
        public string OldName { get; set; }
        public string NewName { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
        public int RecordCount { get; set; }
    }
}
