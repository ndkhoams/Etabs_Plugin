using System;
using System.Collections.Generic;
using System.Linq;
using ETABSv1;

namespace Strip_Rename
{
    public enum SortMode
    {
        LeftToRight_TopToBottom,
        LeftToRight_BottomToTop,
        RightToLeft_TopToBottom,
        RightToLeft_BottomToTop
    }

    public sealed class StripInfo
    {
        public string OldName { get; set; }
        public string TempName { get; set; }
        public string NewName { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
    }

    public static class StripRenamerService
    {
        // Yeu cau chot: startIndex = 1; padding = 3 (001)
        public const int DefaultStartIndex = 1;
        public const int DefaultPadWidth = 3;

        /// <summary>
        /// Tra ve danh sach ten strip hien co trong model (de hien thi len UI cho nguoi dung chon).
        /// </summary>
        public static List<string> GetStripNames(cSapModel sapModel)
        {
            int count = 0;
            string[] names = null;

            int ret = GetStripNameList(sapModel, ref count, ref names);
            if (ret != 0) throw new Exception($"Get strip list failed (ret={ret})");

            return names?.ToList() ?? new List<string>();
        }

        /// <summary>
        /// Tao preview doi ten.
        /// - Chi nhung strip co ten nam trong selectedStripNames moi duoc doi (null = doi toan bo).
        /// - Ten moi luon unique toan model: neu candidate trung voi strip KHONG doi ten
        ///   (hoac trung nhau trong nhom doi ten) thi tu tang index cho den khi khong trung.
        /// </summary>
        public static List<StripInfo> PreviewRename(
            cSapModel sapModel,
            string prefix,
            int startIndex,
            int padWidth,
            SortMode sortMode,
            IEnumerable<string> selectedStripNames = null)
        {
            var allStrips = LoadStripsWithMidpointXY(sapModel);

            HashSet<string> selectedSet = selectedStripNames == null
                ? null
                : new HashSet<string>(selectedStripNames, StringComparer.OrdinalIgnoreCase);

            var toRename = selectedSet == null
                ? allStrips
                : allStrips.Where(s => selectedSet.Contains(s.OldName)).ToList();

            var sorted = SortStrips(toRename, sortMode);

            // Cac ten dang bi chiem boi strip KHONG doi ten -> phai tranh trung.
            var renameSet = new HashSet<string>(sorted.Select(s => s.OldName), StringComparer.OrdinalIgnoreCase);
            var reserved = new HashSet<string>(
                allStrips.Select(s => s.OldName).Where(n => !renameSet.Contains(n)),
                StringComparer.OrdinalIgnoreCase);

            BuildNewNames(sorted, prefix, startIndex, padWidth, reserved);
            ValidateNoDuplicateNewNames(sorted, reserved);

            return sorted;
        }

        public static void ApplyRename(cSapModel sapModel, List<StripInfo> previewRows)
        {
            if (previewRows == null || previewRows.Count == 0) return;

            // Step 1: old -> temp (tranh va cham trung gian)
            foreach (var s in previewRows)
            {
                s.TempName = "__TMP__" + Guid.NewGuid().ToString("N");
                int ret = ChangeStripName(sapModel, s.OldName, s.TempName);
                if (ret != 0) throw new Exception($"Rename failed: {s.OldName} -> {s.TempName} (ret={ret})");
            }

            // Step 2: temp -> new
            foreach (var s in previewRows)
            {
                int ret = ChangeStripName(sapModel, s.TempName, s.NewName);
                if (ret != 0) throw new Exception($"Rename failed: {s.TempName} -> {s.NewName} (ret={ret})");
            }
        }

        private static List<StripInfo> LoadStripsWithMidpointXY(cSapModel sapModel)
        {
            int count = 0;
            string[] names = null;

            int ret = GetStripNameList(sapModel, ref count, ref names);
            if (ret != 0) throw new Exception($"Get strip list failed (ret={ret})");

            var list = new List<StripInfo>();
            if (names == null) return list;

            foreach (var stripName in names)
            {
                int nPts = 0;
                string[] ptNames = null;

                ret = GetStripPoints(sapModel, stripName, ref nPts, ref ptNames);
                if (ret != 0) continue;
                if (ptNames == null || ptNames.Length < 2) continue;

                double x1 = 0, y1 = 0, z1 = 0;
                double x2 = 0, y2 = 0, z2 = 0;

                sapModel.PointObj.GetCoordCartesian(ptNames[0], ref x1, ref y1, ref z1);
                sapModel.PointObj.GetCoordCartesian(ptNames[1], ref x2, ref y2, ref z2);

                list.Add(new StripInfo
                {
                    OldName = stripName,
                    X = (x1 + x2) / 2.0,
                    Y = (y1 + y2) / 2.0
                });
            }

            return list;
        }

        private static List<StripInfo> SortStrips(List<StripInfo> strips, SortMode mode)
        {
            bool xAsc, yAsc;
            switch (mode)
            {
                case SortMode.LeftToRight_TopToBottom: xAsc = true; yAsc = false; break;
                case SortMode.LeftToRight_BottomToTop: xAsc = true; yAsc = true; break;
                case SortMode.RightToLeft_TopToBottom: xAsc = false; yAsc = false; break;
                case SortMode.RightToLeft_BottomToTop: xAsc = false; yAsc = true; break;
                default: xAsc = true; yAsc = false; break;
            }

            var q = xAsc ? strips.OrderBy(s => s.X) : strips.OrderByDescending(s => s.X);
            q = yAsc ? q.ThenBy(s => s.Y) : q.ThenByDescending(s => s.Y);
            return q.ToList();
        }

        private static void BuildNewNames(List<StripInfo> sorted, string prefix, int startIndex, int padWidth, HashSet<string> reserved)
        {
            int i = startIndex;
            var used = new HashSet<string>(
                reserved ?? new HashSet<string>(StringComparer.OrdinalIgnoreCase),
                StringComparer.OrdinalIgnoreCase);

            foreach (var s in sorted)
            {
                string candidate;
                do
                {
                    candidate = $"{prefix}{i.ToString().PadLeft(padWidth, '0')}";
                    i++;
                }
                while (used.Contains(candidate));

                s.NewName = candidate;
                used.Add(candidate);
            }
        }

        private static void ValidateNoDuplicateNewNames(List<StripInfo> rows, HashSet<string> reserved)
        {
            var dup = rows.GroupBy(r => r.NewName, StringComparer.OrdinalIgnoreCase)
                          .FirstOrDefault(g => g.Count() > 1);
            if (dup != null) throw new Exception($"Duplicate new name in preview: {dup.Key}");

            if (reserved != null)
            {
                var clash = rows.FirstOrDefault(r => reserved.Contains(r.NewName));
                if (clash != null)
                    throw new Exception($"New name clashes with an existing strip that is not being renamed: {clash.NewName}");
            }
        }

        // ====== ETABS 22 API mapping ======

        private static int GetStripNameList(cSapModel sapModel, ref int count, ref string[] names)
        {
            return sapModel.StripObj.GetNameList(ref count, ref names);
        }

        private static int GetStripPoints(cSapModel sapModel, string stripName, ref int nPts, ref string[] ptNames)
        {
            return sapModel.StripObj.GetPoints(stripName, ref nPts, ref ptNames);
        }

        private static int ChangeStripName(cSapModel sapModel, string oldName, string newName)
        {
            return sapModel.StripObj.ChangeName(oldName, newName);
        }
    }
}
