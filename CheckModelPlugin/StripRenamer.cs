using ETABSv1;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace Etabs_Ultimate_Tools
{
    /// <summary>
    /// Đọc / đổi tên design strip thông qua Interactive Database Tables của ETABS.
    ///
    /// Lọc theo Group: bảng overwrites nhiều khi KHÔNG lọc được design strip theo tham số
    /// GroupName (trả về tất cả). Vì vậy ta lấy danh sách thành viên của Group qua
    /// cGroup.GetAssignments rồi chỉ giữ các strip có tên nằm trong Group.
    ///
    /// Toạ độ strip: nếu bảng có sẵn cột X/Y thì dùng luôn; nếu chỉ có nhãn điểm
    /// (ví dụ "Strip Object Connectivity") thì gọi PointObj.GetCoordCartesian cho từng điểm.
    /// </summary>
    internal class StripRenamer
    {
        public enum SortMode
        {
            RowLtoR_TtoB = 0,
            RowRtoL_BtoT = 1,
            ColTtoB_LtoR = 2,
            ColBtoT_RtoL = 3
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
        public string CoordDiag = "";

        // Lọc theo Group: tên strip thuộc Group (nếu lấy được) và số strip khớp.
        public HashSet<string> GroupStripNames = null;
        public int GroupFilterCount = -1;

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

        /// <summary>Các bảng có chữ "strip" trong key/tên.</summary>
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

        /// <summary>Danh sách tên Group trong model.</summary>
        public static List<string> GetGroupNames(cSapModel sap)
        {
            var list = new List<string>();
            int n = 0;
            string[] names = null;
            try { sap.GroupDef.GetNameList(ref n, ref names); }
            catch { }
            if (names != null) foreach (var s in names) if (s != null) list.Add(s);
            return list;
        }

        /// <summary>Thành phần được gán vào Group (loại đối tượng + tên).</summary>
        public static void GetGroupAssignments(cSapModel sap, string group, out int[] types, out string[] names)
        {
            int n = 0;
            int[] t = null;
            string[] nm = null;
            try { sap.GroupDef.GetAssignments(group, ref n, ref t, ref nm); }
            catch { }
            types = t ?? new int[0];
            names = nm ?? new string[0];
        }

        /// <summary>Đọc bảng để chỉnh sửa, lọc theo group (qua tham số GroupName của ETABS).</summary>
        public void ReadGroup(cSapModel sap, string tableKey, string groupName)
        {
            TableKey = tableKey;
            GroupName = groupName ?? "";
            GroupStripNames = null;
            GroupFilterCount = -1;
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

        /// <summary>
        /// Lấy danh sách strip thuộc Group qua GetAssignments, chỉ giữ strip có tên khớp.
        /// Trả về: số strip trong bảng khớp Group (>0 = đã lọc); 0 = Group có phần tử nhưng
        /// không khớp strip nào; -1 = không lấy được thành viên Group.
        /// </summary>
        public int ApplyGroupNameFilter(cSapModel sap, string group)
        {
            GroupStripNames = null;
            GroupFilterCount = -1;
            if (NameCol < 0 || string.IsNullOrWhiteSpace(group)) return -1;
            int[] types; string[] names;
            GetGroupAssignments(sap, group, out types, out names);
            if (names == null || names.Length == 0) return -1;

            var set = new HashSet<string>(StringComparer.Ordinal);
            foreach (var s in names) if (!string.IsNullOrWhiteSpace(s)) set.Add(s.Trim());

            var seen = new HashSet<string>(StringComparer.Ordinal);
            int match = 0;
            for (int r = 0; r < NumberRecords; r++)
            {
                string nm = Cell(r, NameCol);
                if (string.IsNullOrWhiteSpace(nm)) continue;
                if (!seen.Add(nm)) continue;
                if (set.Contains(nm)) match++;
            }
            if (match > 0) { GroupStripNames = set; GroupFilterCount = match; return match; }
            return 0;
        }

        /// <summary>
        /// Quét các bảng strip khác để lấy toạ độ tâm mỗi strip. Đọc TOÀN BỘ bảng rồi ghép
        /// theo tên. Ưu tiên cột X/Y; không có thì dùng nhãn điểm + GetCoordCartesian.
        /// </summary>
        public bool LoadCoordinatesFromStripTables(cSapModel sap, string groupName, string excludeTableKey)
        {
            CoordMap = new Dictionary<string, double[]>(StringComparer.Ordinal);
            CoordSource = "";
            var diag = new StringBuilder();
            var candidates = FindStripTables(sap)
                .OrderByDescending(t => ScoreCoordTable(t)).ToList();
            foreach (var t in candidates)
            {
                if (!string.IsNullOrEmpty(excludeTableKey)
                    && string.Equals(t.Key, excludeTableKey, StringComparison.OrdinalIgnoreCase)) continue;
                try
                {
                    string[] fields; string[] data; int nf;
                    if (!ReadDisplay(sap, t.Key, "", out fields, out data, out nf))
                    {
                        diag.AppendLine("• " + t.Key + " → không đọc được (rỗng hoặc lỗi).");
                        continue;
                    }
                    diag.AppendLine("• " + t.Key + " [" + nf + " cột]: " + string.Join(", ", fields));

                    if (TryFillFromXY(fields, data, nf))
                    {
                        CoordSource = t.Key + " (cột X/Y)";
                        CoordDiag = diag.ToString();
                        return true;
                    }
                    if (TryFillFromPoints(sap, fields, data, nf))
                    {
                        CoordSource = t.Key + " (nhãn điểm → GetCoordCartesian)";
                        CoordDiag = diag.ToString();
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    diag.AppendLine("• " + t.Key + " → lỗi: " + ex.Message);
                }
            }
            CoordDiag = diag.ToString();
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

        private static bool ReadDisplay(cSapModel sap, string tableKey, string groupName,
            out string[] fields, out string[] data, out int nf)
        {
            fields = new string[0]; data = new string[0]; nf = 0;
            int tableVersion = 0;
            string[] fieldKeyList = new string[0];
            string[] fieldsIncluded = null;
            int numRecords = 0;
            string[] d = null;
            int ret = sap.DatabaseTables.GetTableForDisplayArray(
                tableKey, ref fieldKeyList, groupName ?? "", ref tableVersion,
                ref fieldsIncluded, ref numRecords, ref d);
            if (ret != 0 || fieldsIncluded == null || d == null) return false;
            fields = fieldsIncluded;
            data = d;
            nf = fieldsIncluded.Length;
            return nf > 0;
        }

        private bool TryFillFromXY(string[] fields, string[] data, int nf)
        {
            int nameC = FindNameCol(fields, true);
            int xC = FindCoordCol(fields, "x");
            int yC = FindCoordCol(fields, "y");
            if (nameC < 0 || xC < 0 || yC < 0) return false;

            var acc = new Dictionary<string, double[]>(StringComparer.Ordinal);
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
            return Finalize(acc);
        }

        private bool TryFillFromPoints(cSapModel sap, string[] fields, string[] data, int nf)
        {
            int nameC = FindNameCol(fields, true);
            if (nameC < 0) return false;
            var pointCols = FindPointCols(fields, nameC);
            if (pointCols.Count == 0) return false;

            var acc = new Dictionary<string, double[]>(StringComparer.Ordinal);
            var cache = new Dictionary<string, double[]>(StringComparer.Ordinal);
            int rows = data.Length / nf;
            for (int r = 0; r < rows; r++)
            {
                string nm = GetCell(data, r, nf, nameC);
                if (string.IsNullOrWhiteSpace(nm)) continue;
                foreach (int pc in pointCols)
                {
                    string pl = GetCell(data, r, nf, pc);
                    if (string.IsNullOrWhiteSpace(pl)) continue;
                    pl = pl.Trim();
                    double[] xy;
                    if (!cache.TryGetValue(pl, out xy))
                    {
                        double px = 0, py = 0, pz = 0;
                        int ret = sap.PointObj.GetCoordCartesian(pl, ref px, ref py, ref pz, "Global");
                        xy = (ret == 0) ? new double[] { px, py } : null;
                        cache[pl] = xy;
                    }
                    if (xy == null) continue;
                    double[] a;
                    if (!acc.TryGetValue(nm, out a)) { a = new double[3]; acc[nm] = a; }
                    a[0] += xy[0];
                    a[1] += xy[1];
                    a[2] += 1;
                }
            }
            return Finalize(acc);
        }

        private bool Finalize(Dictionary<string, double[]> acc)
        {
            if (acc.Count == 0) return false;
            foreach (var kv in acc)
            {
                double c = kv.Value[2] > 0 ? kv.Value[2] : 1;
                CoordMap[kv.Key] = new double[] { kv.Value[0] / c, kv.Value[1] / c };
            }
            return CoordMap.Count > 0;
        }

        private static List<int> FindPointCols(string[] fields, int excludeCol)
        {
            var cols = new List<int>();
            for (int i = 0; i < fields.Length; i++)
            {
                if (i == excludeCol) continue;
                string fl = (fields[i] ?? "").ToLowerInvariant().Replace(" ", "");
                if (fl.Contains("number") || fl.Contains("count") || fl.Contains("num")) continue;
                if (fl.Contains("point") || fl.Contains("joint")) cols.Add(i);
            }
            return cols;
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

        /// <summary>Gom record theo tên strip (lọc theo Group nếu có), sắp theo vị trí, gán tên mới.</summary>
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
                if (GroupStripNames != null && GroupStripNames.Count > 0 && !GroupStripNames.Contains(oldName)) continue;
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
