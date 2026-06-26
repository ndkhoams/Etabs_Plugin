using ETABSv1;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Etabs_Ultimate_Tools
{
    /// <summary>Khả năng chịu tải của một loại cọc (point spring).</summary>
    public class PileSpringType
    {
        public string Name { get; set; } = "";
        public double TensionCap { get; set; }      // SCT chịu kéo (kN)
        public double CompressionCap { get; set; }  // SCT chịu nén (kN)
    }

    /// <summary>Một dòng kết quả kiểm tra phản lực cọc.</summary>
    public class PileReactionRow
    {
        public string PileType { get; set; } = "";   // Loại cọc
        public string PileId { get; set; } = "";      // Số hiệu cọc (label điểm)
        public string Combo { get; set; } = "";        // Tổ hợp
        public double Reaction { get; set; }           // Phản lực đầu cọc (kN): + nén / - kéo
        public double TensionCap { get; set; }         // SCT chịu kéo (kN)
        public double CompressionCap { get; set; }     // SCT chịu nén (kN)
        public string Result { get; set; } = "";       // Kết luận
    }

    /// <summary>Một trường hợp tải -> một sheet trong file Excel.</summary>
    public class PileReactionCase
    {
        public string Title { get; set; } = "";
        public string SheetName { get; set; } = "";
        public string Combo { get; set; } = "";
        public List<PileReactionRow> Rows { get; set; } = new List<PileReactionRow>();
    }

    /// <summary>
    /// Phát hiện các điểm có gán point spring (= cọc), đọc phản lực đầu cọc (F3)
    /// theo từng tổ hợp và so sánh với SCT chịu kéo/nén.
    /// Port net48 (không dùng record/init/MaxBy/range operator).
    /// </summary>
    public static class PileReactionChecker
    {
        private class PilePoint
        {
            public string Name = "";
            public string Label = "";
            public double Kz = 0.0;   // độ cứng lò xo phương đứng (U3)
        }

        // ── Liệt kê các loại point spring khai báo trong model (= loại cọc) ──────
        public static List<string> GetSpringTypes(cSapModel sap)
        {
            int n = 0;
            string[] names = null;
            try { sap.PropPointSpring.GetNameList(ref n, ref names); }
            catch { return new List<string>(); }

            if (names == null) return new List<string>();
            return names.Where(s => !string.IsNullOrWhiteSpace(s))
                        .Select(s => s.Trim())
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
                        .ToList();
        }

        // ── Các "loại cọc" thực sự phát hiện được trên model (để đồng bộ bảng SCT) ─
        public static List<string> GetPileTypeKeys(cSapModel sap)
        {
            var piles = GetPilePoints(sap);
            var defined = GetSpringTypes(sap);
            return piles.Select(p => ResolveType(p, defined))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
                        .ToList();
        }

        // ── Tính phản lực + kết luận cho 1 tổ hợp tải ────────────────────────────
        public static List<PileReactionRow> Compute(cSapModel sap, string combo,
            Dictionary<string, PileSpringType> caps)
        {
            var rows = new List<PileReactionRow>();
            if (string.IsNullOrWhiteSpace(combo)) return rows;

            sap.SetPresentUnits(eUnits.kN_m_C);

            var piles = GetPilePoints(sap);
            if (piles.Count == 0) return rows;

            var defined = GetSpringTypes(sap);
            EtabsHelper.SelectCaseOrCombo(sap, combo);

            foreach (var pile in piles)
            {
                double pmax, pmin;
                if (!TryGetReactionRange(sap, pile.Name, out pmax, out pmin))
                    continue;

                string type = ResolveType(pile, defined);

                double tensCap = 0, compCap = 0;
                if (caps != null)
                {
                    PileSpringType cap;
                    if (caps.TryGetValue(type, out cap))
                    {
                        tensCap = cap.TensionCap;
                        compCap = cap.CompressionCap;
                    }
                }

                double comp = pmax > 0 ? pmax : 0.0;   // phần nén  (phản lực dương)
                double tens = pmin < 0 ? -pmin : 0.0;  // phần kéo  (phản lực âm)

                bool hasComp = compCap > 0;
                bool hasTens = tensCap > 0;
                bool okComp = !hasComp || comp <= compCap;
                bool okTens = !hasTens || tens <= tensCap;

                string result;
                if (!hasComp && !hasTens)
                    result = "Chưa nhập SCT";
                else
                    result = (okComp && okTens) ? "Đạt" : "Không Đạt";

                double compUtil = hasComp ? comp / compCap : 0.0;
                double tensUtil = hasTens ? tens / tensCap : 0.0;
                double reaction = compUtil >= tensUtil ? pmax : pmin;

                rows.Add(new PileReactionRow
                {
                    PileType = type,
                    PileId = pile.Label,
                    Combo = combo,
                    Reaction = reaction,
                    TensionCap = tensCap,
                    CompressionCap = compCap,
                    Result = result
                });
            }

            return rows
                .OrderBy(r => r.PileType, StringComparer.OrdinalIgnoreCase)
                .ThenBy(r => r.PileId, new NaturalComparer())
                .ToList();
        }

        // Loại cọc: nếu chỉ có 1 loại point spring khai báo -> dùng tên đó cho mọi cọc;
        // nếu chưa khai báo -> "Cọc"; nếu nhiều loại -> nhóm theo độ cứng đứng Kz.
        private static string ResolveType(PilePoint p, List<string> defined)
        {
            if (defined != null && defined.Count == 1) return defined[0];
            if (defined == null || defined.Count == 0) return "Cọc";
            return "Kz=" + p.Kz.ToString("0", CultureInfo.InvariantCulture);
        }

        // ── Lấy danh sách điểm có gán point spring (= cọc) ──────────────────────
        private static List<PilePoint> GetPilePoints(cSapModel sap)
        {
            var list = new List<PilePoint>();

            int n = 0;
            string[] names = null;
            try
            {
                if (sap.PointObj.GetNameList(ref n, ref names) != 0 || names == null)
                    return list;
            }
            catch { return list; }

            foreach (string pt in names)
            {
                if (string.IsNullOrWhiteSpace(pt)) continue;

                double kz;
                if (!HasSpring(sap, pt, out kz)) continue;

                string label = pt, story = "";
                try { sap.PointObj.GetLabelFromName(pt, ref label, ref story); }
                catch { label = pt; }

                list.Add(new PilePoint
                {
                    Name = pt,
                    Label = string.IsNullOrWhiteSpace(label) ? pt : label.Trim(),
                    Kz = kz
                });
            }

            return list;
        }

        // GetSpring trả về 0 KHI VÀ CHỈ KHI điểm có lò xo (tổng hợp mọi spring gán
        // cho điểm, kể cả gán qua point spring property). K là 6 số hạng đường chéo.
        private static bool HasSpring(cSapModel sap, string pointName, out double kz)
        {
            kz = 0.0;
            try
            {
                double[] k = new double[6];
                int ret = sap.PointObj.GetSpring(pointName, ref k);
                if (ret != 0 || k == null || k.Length < 3) return false;

                bool any = false;
                for (int i = 0; i < k.Length; i++)
                    if (Math.Abs(k[i]) > 1e-12) { any = true; break; }
                if (!any) return false;

                kz = Math.Abs(k[2]);   // U3 (phương đứng)
                return true;
            }
            catch { return false; }
        }

        // ── Đọc phản lực F3 (phương đứng) của 1 điểm: max & min trên mọi step ────
        private static bool TryGetReactionRange(cSapModel sap, string pointName,
            out double pmax, out double pmin)
        {
            pmax = double.MinValue;
            pmin = double.MaxValue;

            int num = 0;
            string[] obj = new string[0], elm = new string[0];
            string[] lc = new string[0], stepType = new string[0];
            double[] stepNum = new double[0];
            double[] f1 = new double[0], f2 = new double[0], f3 = new double[0];
            double[] m1 = new double[0], m2 = new double[0], m3 = new double[0];

            int ret;
            try
            {
                ret = sap.Results.JointReact(pointName, eItemTypeElm.ObjectElm,
                    ref num, ref obj, ref elm, ref lc, ref stepType, ref stepNum,
                    ref f1, ref f2, ref f3, ref m1, ref m2, ref m3);
            }
            catch { return false; }

            if (ret != 0 || num == 0 || f3 == null || f3.Length == 0)
                return false;

            int count = Math.Min(num, f3.Length);
            for (int i = 0; i < count; i++)
            {
                if (f3[i] > pmax) pmax = f3[i];
                if (f3[i] < pmin) pmin = f3[i];
            }

            return pmax != double.MinValue && pmin != double.MaxValue;
        }

        // So sánh "tự nhiên" để 2,10,100 sắp đúng thứ tự số.
        private class NaturalComparer : IComparer<string>
        {
            public int Compare(string a, string b)
            {
                long na, nb;
                bool ia = long.TryParse((a ?? "").Trim(), out na);
                bool ib = long.TryParse((b ?? "").Trim(), out nb);
                if (ia && ib) return na.CompareTo(nb);
                return string.Compare(a, b, StringComparison.OrdinalIgnoreCase);
            }
        }
    }
}
