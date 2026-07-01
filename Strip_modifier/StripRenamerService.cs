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
        public static List<StripInfo> PreviewRename(
            cSapModel sapModel,
            string prefix,
            int startIndex,
            int padWidth,
            SortMode sortMode)
        {
            var strips = LoadStripsWithMidpointXY(sapModel);
            var sorted = SortStrips(strips, sortMode);

            BuildNewNames(sorted, prefix, startIndex, padWidth);
            ValidateNoDuplicateNewNames(sorted);

            return sorted;
        }

        public static void ApplyRename(cSapModel sapModel, List<StripInfo> previewRows)
        {
            if (previewRows == null || previewRows.Count == 0) return;

            // Step 1: old -> temp
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

        private static void BuildNewNames(List<StripInfo> sorted, string prefix, int startIndex, int padWidth)
        {
            int i = startIndex;
            foreach (var s in sorted)
            {
                s.NewName = $"{prefix}{i.ToString().PadLeft(padWidth, '0')}";
                i++;
            }
        }

        private static void ValidateNoDuplicateNewNames(List<StripInfo> rows)
        {
            var dup = rows.GroupBy(r => r.NewName, StringComparer.OrdinalIgnoreCase)
                          .FirstOrDefault(g => g.Count() > 1);
            if (dup != null) throw new Exception($"Duplicate new name in preview: {dup.Key}");
        }

        // ====== ETABS 22 API mapping (bạn map 3 hàm này theo IntelliSense) ======

        private static int GetStripNameList(cSapModel sapModel, ref int count, ref string[] names)
        {
            // ETABS 22: map đúng hàm thực tế ở đây:
            // return sapModel.StripObj.GetNameList(ref count, ref names);
            throw new NotImplementedException("Map StripObj.GetNameList(...) theo ETABS 22.");
        }

        private static int GetStripPoints(cSapModel sapModel, string stripName, ref int nPts, ref string[] ptNames)
        {
            // ETABS 22: map đúng hàm thực tế ở đây:
            // return sapModel.StripObj.GetPoints(stripName, ref nPts, ref ptNames);
            throw new NotImplementedException("Map StripObj.GetPoints(...) theo ETABS 22.");
        }

        private static int ChangeStripName(cSapModel sapModel, string oldName, string newName)
        {
            // ETABS 22: map đúng hàm rename thực tế:
            // return sapModel.StripObj.ChangeName(oldName, newName);
            throw new NotImplementedException("Map StripObj.ChangeName(...) theo ETABS 22.");
        }
    }
}