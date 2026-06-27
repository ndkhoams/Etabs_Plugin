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
        public double HorizontalCap { get; set; }   // SCT chịu ngang (kN)
    }

    /// <summary>Thông tin 1 loại cọc phát hiện trên model (để đồng bộ + gợi ý SCT).</summary>
    public class PileTypeInfo
    {
        public string Key { get; set; } = "";
        public double Kz { get; set; }          // độ cứng lò xo phương đứng (kN/m)
        public double DefaultCap { get; set; }  // SCT tạm = Kz * 0.01 (kN)
    }

    /// <summary>Một dòng kết quả kiểm tra phản lực cọc.</summary>
    public class PileReactionRow
    {
        public string PileType { get; set; } = "";   // Loại cọc
        public string PileId { get; set; } = "";      // Số hiệu cọc (label điểm)
        public string Combo { get; set; } = "";        // Tổ hợp (có thể kèm _max/_min)
        public double Reaction { get; set; }           // Phản lực đầu cọc (kN): + nén / - kéo
        public double TensionCap { get; set; }         // SCT chịu kéo (kN)
        public double CompressionCap { get; set; }     // SCT chịu nén (kN)
        public string Result { get; set; } = "";       // Kết luận (đứng)
        public double Fx { get; set; }                 // |FX| tại step có H lớn nhất (kN)
        public double Fy { get; set; }                 // |FY| tại step có H lớn nhất (kN)
        public double Horizontal { get; set; }         // Hợp lực ngang H = sqrt(FX^2+FY^2) (kN)
        public double HorizontalCap { get; set; }      // SCT chịu ngang (kN)
        public string HResult { get; set; } = "";      // Kết luận ngang
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
    /// Tổ hợp bao (Max/Min) -> 2 dòng/cọc trong cùng 1 bảng (tên tổ hợp kèm _max/_min).
    /// Port net48 (không dùng record/init/MaxBy/range operator).
    /// </summary>
    public static class PileReactionChecker
    {
        private const double SctFactor = 0.01; // SCT tạm = Kz (kN/m) * 0.01 (m)

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

        // ── Các loại cọc thực phát hiện trên model + SCT gợi ý (Kz * 0.01) ─────
        public static List<PileTypeInfo> GetPileTypeInfos(cSapModel sap)
        {
            var piles = GetPilePoints(sap);
            var defined = GetSpringTypes(sap);

            var maxKz = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            foreach (var p in piles)
            {
                string key = ResolveType(p, defined);
                double cur;
                if (!maxKz.TryGetValue(key, out cur) || p.Kz > cur) maxKz[key] = p.Kz;
            }

            return maxKz
                .OrderBy(kv => kv.Key, StringComparer.OrdinalIgnoreCase)
                .Select(kv => new PileTypeInfo
                {
                    Key = kv.Key,
                    Kz = kv.Value,
                    DefaultCap = kv.Value * SctFactor
                })
                .ToList();
        }

        // ── Tính phản lực + kết luận cho 1 tổ hợp tải (1 case = 1 sheet) ─────────────
        // Tổ hợp bao -> mỗi cọc 2 dòng: tên tổ hợp + "_max" (nén) và + "_min" (kéo).
        public static PileReactionCase ComputeCase(cSapModel sap, string combo,
            string title, string sheet, Dictionary<string, PileSpringType> caps)
        {
            if (string.IsNullOrWhiteSpace(combo)) return null;

            sap.SetPresentUnits(eUnits.kN_m_C);

            // Bật chế độ bao (Envelopes): tổ hợp bao trả về CẢ Max và Min cho mỗi điểm.
            try { sap.Results.Setup.SetOptionMultiValuedCombo(1); } catch { }

            var piles = GetPilePoints(sap);
            if (piles.Count == 0) return null;

            var defined = GetSpringTypes(sap);
            EtabsHelper.SelectCaseOrCombo(sap, combo);

            var rows = new List<PileReactionRow>();
            foreach (var pile in piles)
            {
                double pmax, pmin;
                bool multi;
                if (!TryGetReactionRange(sap, pile.Name, out pmax, out pmin, out multi))
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

                if (multi)
                {
                    rows.Add(BuildCompRow(type, pile.Label, combo + "_max", pmax, tensCap, compCap));
                    rows.Add(BuildTensRow(type, pile.Label, combo + "_min", pmin, tensCap, compCap));
                }
                else
                {
                    rows.Add(BuildSingleRow(type, pile.Label, combo, pmax, tensCap, compCap, true));
                }
            }

            if (rows.Count == 0) return null;

            var sorted = rows
                .OrderBy(r => r.PileType, StringComparer.OrdinalIgnoreCase)
                .ThenBy(r => r.PileId, new NaturalComparer())
                .ThenBy(r => r.Combo, StringComparer.OrdinalIgnoreCase)
                .ToList();

            return new PileReactionCase { Title = title, SheetName = sheet, Combo = combo, Rows = sorted };
        }

        // ── Tính phản lực đứng (F3) + hợp lực ngang H = sqrt(FX^2+FY^2) cho 1 tổ hợp ─────
        // H tính CHÍNH XÁC theo từng bước: tắt chế độ bao (envelope), duyệt mọi step,
        // tính H tại từng step rồi lấy step có H lớn nhất (FX, FY tương quan cùng 1 bước).
        // considerTension = false -> bỏ qua kiểm tra SCT kéo (không sinh dòng _min, dòng đơn chỉ xét nén).
        public static PileReactionCase ComputeCaseH(cSapModel sap, string combo,
            string title, string sheet, Dictionary<string, PileSpringType> caps, bool considerTension)
        {
            if (string.IsNullOrWhiteSpace(combo)) return null;

            sap.SetPresentUnits(eUnits.kN_m_C);
            // 0 = trả về từng step (không bao) để lấy được cặp FX-FY tương quan cùng 1 bước.
            try { sap.Results.Setup.SetOptionMultiValuedCombo(0); } catch { }

            var piles = GetPilePoints(sap);
            if (piles.Count == 0) return null;

            var defined = GetSpringTypes(sap);
            EtabsHelper.SelectCaseOrCombo(sap, combo);

            var rows = new List<PileReactionRow>();
            foreach (var pile in piles)
            {
                double pmax, pmin, fxAbs, fyAbs;
                bool multi;
                if (!TryGetReactionRangeH(sap, pile.Name, out pmax, out pmin, out fxAbs, out fyAbs, out multi))
                    continue;

                double h = Math.Sqrt(fxAbs * fxAbs + fyAbs * fyAbs);

                string type = ResolveType(pile, defined);
                double tensCap = 0, compCap = 0, horizCap = 0;
                if (caps != null)
                {
                    PileSpringType cap;
                    if (caps.TryGetValue(type, out cap))
                    {
                        tensCap = cap.TensionCap;
                        compCap = cap.CompressionCap;
                        horizCap = cap.HorizontalCap;
                    }
                }

                string hResult = horizCap <= 0 ? "Chưa nhập SCT" : (h <= horizCap ? "Đạt" : "Không Đạt");

                if (multi)
                {
                    var rc = BuildCompRow(type, pile.Label, combo + "_max", pmax, tensCap, compCap);
                    FillH(rc, fxAbs, fyAbs, h, horizCap, hResult);
                    rows.Add(rc);
                    if (considerTension)
                    {
                        var rt = BuildTensRow(type, pile.Label, combo + "_min", pmin, tensCap, compCap);
                        FillH(rt, fxAbs, fyAbs, h, horizCap, hResult);
                        rows.Add(rt);
                    }
                }
                else
                {
                    var rs = BuildSingleRow(type, pile.Label, combo, pmax, tensCap, compCap, considerTension);
                    FillH(rs, fxAbs, fyAbs, h, horizCap, hResult);
                    rows.Add(rs);
                }
            }

            if (rows.Count == 0) return null;

            var sorted = rows
                .OrderBy(r => r.PileType, StringComparer.OrdinalIgnoreCase)
                .ThenBy(r => r.PileId, new NaturalComparer())
                .ThenBy(r => r.Combo, StringComparer.OrdinalIgnoreCase)
                .ToList();

            return new PileReactionCase { Title = title, SheetName = sheet, Combo = combo, Rows = sorted };
        }

        private static void FillH(PileReactionRow r, double fxAbs, double fyAbs, double h,
            double horizCap, string hResult)
        {
            r.Fx = fxAbs;
            r.Fy = fyAbs;
            r.Horizontal = h;
            r.HorizontalCap = horizCap;
            r.HResult = hResult;
        }

        // Dòng kiểm tra NÉN (dùng Max): phản lực dương lớn nhất so với SCT nén.
        private static PileReactionRow BuildCompRow(string type, string id, string comboLabel,
            double pmax, double tensCap, double compCap)
        {
            double comp = pmax > 0 ? pmax : 0.0;
            string result;
            if (compCap <= 0) result = "Chưa nhập SCT";
            else result = comp <= compCap ? "Đạt" : "Không Đạt";
            return new PileReactionRow
            {
                PileType = type, PileId = id, Combo = comboLabel, Reaction = pmax,
                TensionCap = tensCap, CompressionCap = compCap, Result = result
            };
        }

        // Dòng kiểm tra KÉO (dùng Min): phản lực âm nhỏ nhất so với SCT kéo.
        private static PileReactionRow BuildTensRow(string type, string id, string comboLabel,
            double pmin, double tensCap, double compCap)
        {
            double tens = pmin < 0 ? -pmin : 0.0;
            string result;
            if (tensCap <= 0) result = "Chưa nhập SCT";
            else result = tens <= tensCap ? "Đạt" : "Không Đạt";
            return new PileReactionRow
            {
                PileType = type, PileId = id, Combo = comboLabel, Reaction = pmin,
                TensionCap = tensCap, CompressionCap = compCap, Result = result
            };
        }

        // Dòng tổ hợp 1 giá trị: so cả kéo lẫn nén trên cùng 1 dòng.
        // considerTension = false -> chỉ xét nén (bỏ qua kéo).
        private static PileReactionRow BuildSingleRow(string type, string id, string combo,
            double v, double tensCap, double compCap, bool considerTension)
        {
            double comp = v > 0 ? v : 0.0;
            double tens = v < 0 ? -v : 0.0;
            bool hasComp = compCap > 0, hasTens = considerTension && tensCap > 0;
            bool okComp = !hasComp || comp <= compCap;
            bool okTens = !hasTens || tens <= tensCap;
            string result = (!hasComp && !hasTens)
                ? "Chưa nhập SCT"
                : ((okComp && okTens) ? "Đạt" : "Không Đạt");
            return new PileReactionRow
            {
                PileType = type, PileId = id, Combo = combo, Reaction = v,
                TensionCap = tensCap, CompressionCap = compCap, Result = result
            };
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
        // multi = true nếu tổ hợp trả về > 1 giá trị (tổ hợp bao Max/Min).
        private static bool TryGetReactionRange(cSapModel sap, string pointName,
            out double pmax, out double pmin, out bool multi)
        {
            pmax = double.MinValue;
            pmin = double.MaxValue;
            multi = false;

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

            multi = count > 1;
            return pmax != double.MinValue && pmin != double.MaxValue;
        }

        // ── Như TryGetReactionRange nhưng lấy thêm FX, FY tại step có H lớn nhất ──────────
        // Duyệt từng step: tính H = sqrt(F1^2 + F2^2) tại CHÍNH step đó (FX, FY tương quan)
        // rồi giữ lại FX, FY của step có H lớn nhất.
        private static bool TryGetReactionRangeH(cSapModel sap, string pointName,
            out double pmax, out double pmin, out double fxAbs, out double fyAbs, out bool multi)
        {
            pmax = double.MinValue;
            pmin = double.MaxValue;
            fxAbs = 0.0;
            fyAbs = 0.0;
            multi = false;

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
            double maxH = -1.0;
            for (int i = 0; i < count; i++)
            {
                if (f3[i] > pmax) pmax = f3[i];
                if (f3[i] < pmin) pmin = f3[i];

                // Hợp lực ngang tại CHÍNH step này (FX, FY tương quan cùng 1 bước).
                double fx = (f1 != null && i < f1.Length) ? f1[i] : 0.0;
                double fy = (f2 != null && i < f2.Length) ? f2[i] : 0.0;
                double hStep = Math.Sqrt(fx * fx + fy * fy);
                if (hStep > maxH)
                {
                    maxH = hStep;
                    fxAbs = Math.Abs(fx);
                    fyAbs = Math.Abs(fy);
                }
            }

            multi = count > 1;
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
