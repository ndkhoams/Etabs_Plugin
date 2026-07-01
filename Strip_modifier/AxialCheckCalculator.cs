using ETABSv1;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Etabs_Ultimate_Tools
{
    /// <summary>
    /// Tính toán kiểm tra hệ số lực dọc quy đổi cho cột và vách Pier.
    /// Tách hoàn toàn khỏi UI để tái sử dụng trong tab "Check lực dọc".
    /// Port về net48 (không dùng record/init/MaxBy/HashCode/range operator).
    /// </summary>
    public sealed class AxialCheckCalculator
    {
        private readonly cSapModel _model;
        private readonly double _columnLimit;
        private readonly double _wallLimit;
        private readonly double _fckCube;
        private readonly double _fck;
        private readonly double _fcd;

        private Dictionary<string, SectionInfo> _frameSectionCache;
        private Dictionary<string, SectionInfo> _pierSectionCache;
        private Dictionary<string, double> _storyElevCache;

        private sealed class SectionInfo
        {
            public double Ac;
            public double T3;
            public double T2;
            public string Material = "";
            public double FckCube;
            public double Fck;
            public double Fcd;
        }

        private sealed class PierStoryKey
        {
            public string Pier;
            public string Story;
            public PierStoryKey(string pier, string story) { Pier = pier; Story = story; }
        }

        public AxialCheckCalculator(cSapModel model, double fckCube,
            double alphaCc, double gammaC, double columnLimit, double wallLimit)
        {
            _model = model;
            _fckCube = fckCube;
            _columnLimit = columnLimit;
            _wallLimit = wallLimit;

            _fck = Math.Floor(fckCube * 0.8);
            _fcd = Math.Floor(alphaCc * _fck / gammaC);
        }

        // ── Build toàn bộ rows ────────────────────────────────────────────────────
        public List<AxialCheckRow> Build(string combo)
        {
            if (string.IsNullOrWhiteSpace(combo))
                throw new InvalidOperationException("Chưa chọn combo kiểm tra.");
            if (_fckCube <= 0 || _fck <= 0 || _fcd <= 0)
                throw new InvalidOperationException("Cấp bền bê tông không hợp lệ.");

            _frameSectionCache = new Dictionary<string, SectionInfo>(StringComparer.OrdinalIgnoreCase);
            _pierSectionCache = new Dictionary<string, SectionInfo>(StringComparer.OrdinalIgnoreCase);
            _storyElevCache = null;

            HashSet<string> frameNames;
            List<PierStoryKey> pierStoryKeys;
            ReadSelection(out frameNames, out pierStoryKeys);

            SelectOnlyCombo(combo);

            var rows = new List<AxialCheckRow>();
            int stt = 1;

            foreach (string frame in frameNames)
                AddFrameRows(frame, combo, rows, ref stt);

            if (pierStoryKeys.Count > 0)
                AddPierRows(pierStoryKeys, combo, rows, ref stt);

            if (rows.Count == 0)
                throw new InvalidOperationException(
                    "Không có dữ liệu để xuất.\n" +
                    "Hãy chọn trực tiếp cột hoặc area vách Pier trong ETABS và chạy Analyze trước.");

            rows = SortByStory(rows);
            for (int i = 0; i < rows.Count; i++) rows[i].STT = i + 1;
            return rows;
        }

        // ── Đọc selection từ ETABS (1 lần duy nhất) ─────────────────────────────
        private void ReadSelection(out HashSet<string> frameNames, out List<PierStoryKey> pierStoryKeys)
        {
            int count = 0;
            int[] types = new int[0];
            string[] names = new string[0];

            _model.SelectObj.GetSelected(ref count, ref types, ref names);

            if (count == 0)
                throw new InvalidOperationException("Chưa chọn cột hoặc vách trong ETABS.");

            frameNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            pierStoryKeys = new List<PierStoryKey>();
            var seenPierKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < count; i++)
            {
                string objName = names[i];

                if (types[i] == 2) // Frame
                {
                    string pierName = "";
                    bool hasPier = _model.FrameObj.GetPier(objName, ref pierName) == 0
                                   && !string.IsNullOrWhiteSpace(pierName)
                                   && !pierName.Equals("None", StringComparison.OrdinalIgnoreCase);

                    if (hasPier)
                    {
                        string story = "", label = objName;
                        _model.FrameObj.GetLabelFromName(objName, ref label, ref story);
                        AddPierStoryKey(pierName, story, pierStoryKeys, seenPierKeys);
                    }
                    else
                    {
                        frameNames.Add(objName);
                    }
                }
                else if (types[i] == 5) // Area
                {
                    string pierName = "";
                    if (_model.AreaObj.GetPier(objName, ref pierName) != 0
                        || string.IsNullOrWhiteSpace(pierName)
                        || pierName.Equals("None", StringComparison.OrdinalIgnoreCase))
                        continue;

                    string story = "", label = objName;
                    _model.AreaObj.GetLabelFromName(objName, ref label, ref story);
                    AddPierStoryKey(pierName, story, pierStoryKeys, seenPierKeys);
                }
            }
        }

        private static void AddPierStoryKey(string pierName, string story,
            List<PierStoryKey> list, HashSet<string> seen)
        {
            if (string.IsNullOrWhiteSpace(story)) return;
            string key = NormalizePier(pierName) + "|" + NormalizeStory(story);
            if (seen.Add(key))
                list.Add(new PierStoryKey(pierName.Trim(), story.Trim()));
        }

        // ── Frame Column rows ─────────────────────────────────────────────────────
        private void AddFrameRows(string frameName, string combo,
            List<AxialCheckRow> rows, ref int stt)
        {
            eFrameDesignOrientation orient = eFrameDesignOrientation.Null;
            if (_model.FrameObj.GetDesignOrientation(frameName, ref orient) != 0
                || orient != eFrameDesignOrientation.Column)
                return;

            string propName = "", sAuto = "";
            if (_model.FrameObj.GetSection(frameName, ref propName, ref sAuto) != 0
                || string.IsNullOrWhiteSpace(propName))
                return;

            var sec = GetFrameSection(propName);
            if (sec == null) return;

            ApplyConcreteToSection(sec);

            string story = "", label = frameName;
            _model.FrameObj.GetLabelFromName(frameName, ref label, ref story);

            double ned = GetFrameMaxAxial(frameName);
            if (double.IsNaN(ned)) return;

            double acFcd = sec.Ac * sec.Fcd * 1000.0;
            if (acFcd <= 0) return;

            double nud = Math.Abs(ned) / acFcd;

            rows.Add(MakeRow(ref stt, story, label, "Column", combo, sec, ned, acFcd, nud, _columnLimit));
        }

        // ── Pier rows (gọi PierForce một lần cho tất cả pier) ────────────────────
        private void AddPierRows(IList<PierStoryKey> pierStoryKeys, string combo,
            List<AxialCheckRow> rows, ref int stt)
        {
            int count = 0;
            string[] storyArr = new string[0], pierArr = new string[0];
            string[] caseArr = new string[0], locArr = new string[0];
            double[] p = new double[0], v2 = new double[0], v3 = new double[0],
                     t = new double[0], m2 = new double[0], m3 = new double[0];

            if (_model.Results.PierForce(
                    ref count, ref storyArr, ref pierArr,
                    ref caseArr, ref locArr,
                    ref p, ref v2, ref v3, ref t, ref m2, ref m3) != 0 || count == 0)
                return;

            // Build lookup: "normPier|normStory" -> index của |P| lớn nhất
            var lookupIdx = new Dictionary<string, int>(StringComparer.Ordinal);
            var lookupMax = new Dictionary<string, double>(StringComparer.Ordinal);
            for (int k = 0; k < count; k++)
            {
                string key = NormalizePier(pierArr[k]) + "|" + NormalizeStory(storyArr[k]);
                double absP = Math.Abs(p[k]);
                double cur;
                if (!lookupMax.TryGetValue(key, out cur) || absP > cur)
                {
                    lookupMax[key] = absP;
                    lookupIdx[key] = k;
                }
            }

            foreach (var psk in pierStoryKeys)
            {
                string key = NormalizePier(psk.Pier) + "|" + NormalizeStory(psk.Story);
                int idx;
                if (!lookupIdx.TryGetValue(key, out idx)) continue;

                double ned = p[idx];

                var sec = GetPierSection(psk.Pier, psk.Story);
                if (sec == null) continue;

                ApplyConcreteToSection(sec);

                double ac = sec.T2 * sec.T3;
                double acFcd = ac * sec.Fcd * 1000.0;
                if (acFcd <= 0 || double.IsInfinity(acFcd)) continue;

                double vd = Math.Abs(ned) / acFcd;

                double ratio = sec.T3 > 0 ? Math.Max(sec.T2, sec.T3) / Math.Min(sec.T2, sec.T3) : 0;
                bool isWall = ratio > 4.0;
                double limit = isWall ? _wallLimit : _columnLimit;
                string type = isWall ? "Pier.Wall" : "Pier.Col";

                rows.Add(MakeRow(ref stt, psk.Story, psk.Pier, type, combo, sec, ned, acFcd, vd, limit));
            }
        }

        // ── Lấy lực dọc lớn nhất của một frame ───────────────────────────────────
        private double GetFrameMaxAxial(string frameName)
        {
            int count = 0;
            string[] obj = new string[0], elm = new string[0];
            string[] cas = new string[0], sType = new string[0];
            double[] sta = new double[0], eSta = new double[0];
            double[] sNum = new double[0], p = new double[0];
            double[] v2 = new double[0], v3 = new double[0];
            double[] tT = new double[0], m2 = new double[0];
            double[] m3 = new double[0];

            int ret = _model.Results.FrameForce(
                frameName, eItemTypeElm.ObjectElm,
                ref count, ref obj, ref sta, ref elm, ref eSta,
                ref cas, ref sType, ref sNum,
                ref p, ref v2, ref v3, ref tT, ref m2, ref m3);

            if (ret != 0 || count == 0 || p.Length == 0) return double.NaN;

            double best = p[0];
            for (int i = 1; i < p.Length; i++)
                if (Math.Abs(p[i]) > Math.Abs(best)) best = p[i];
            return best;
        }

        // ── Section cache ─────────────────────────────────────────────────────────
        private SectionInfo GetFrameSection(string propName)
        {
            if (_frameSectionCache == null)
                _frameSectionCache = new Dictionary<string, SectionInfo>(StringComparer.OrdinalIgnoreCase);

            SectionInfo cached;
            if (!_frameSectionCache.TryGetValue(propName, out cached))
            {
                cached = FetchFrameSection(propName);
                _frameSectionCache[propName] = cached;
            }
            return cached;
        }

        private SectionInfo GetPierSection(string pierName, string story)
        {
            if (_pierSectionCache == null)
                _pierSectionCache = new Dictionary<string, SectionInfo>(StringComparer.OrdinalIgnoreCase);

            string key = NormalizePier(pierName) + "|" + NormalizeStory(story);
            SectionInfo cached;
            if (!_pierSectionCache.TryGetValue(key, out cached))
            {
                cached = FetchPierSection(pierName, story);
                _pierSectionCache[key] = cached;
            }
            return cached;
        }

        // ── Lấy tiết diện Frame từ API ────────────────────────────────────────────
        private SectionInfo FetchFrameSection(string propName)
        {
            string f = "", mat = "", n = "", g = "";
            int c = 0;
            double t3 = 0, t2 = 0;

            if (_model.PropFrame.GetRectangle(propName, ref f, ref mat, ref t3, ref t2, ref c, ref n, ref g) == 0
                && t3 > 0 && t2 > 0)
                return new SectionInfo { Ac = t3 * t2, T3 = t3, T2 = t2, Material = mat };

            double dia = 0; mat = ""; f = ""; n = ""; g = "";
            if (_model.PropFrame.GetCircle(propName, ref f, ref mat, ref dia, ref c, ref n, ref g) == 0
                && dia > 0)
                return new SectionInfo { Ac = Math.PI * dia * dia / 4.0, T3 = dia, T2 = dia, Material = mat };

            double area = 0, as2 = 0, as3 = 0, tor = 0, i22 = 0, i33 = 0;
            double s22 = 0, s33 = 0, z22 = 0, z33 = 0, r22 = 0, r33 = 0;
            if (_model.PropFrame.GetSectProps(propName,
                    ref area, ref as2, ref as3, ref tor, ref i22, ref i33,
                    ref s22, ref s33, ref z22, ref z33, ref r22, ref r33) == 0 && area > 0)
            {
                mat = TryGetFrameMaterial(propName);
                return new SectionInfo { Ac = area, T3 = 0, T2 = 0, Material = mat };
            }

            return null;
        }

        // Lấy material bằng reflection để né phụ thuộc 'dynamic' (Microsoft.CSharp) trên net48.
        private string TryGetFrameMaterial(string propName)
        {
            try
            {
                object pf = _model.PropFrame;
                var method = pf.GetType().GetMethod("GetMaterial",
                    new[] { typeof(string), typeof(string).MakeByRefType() });
                if (method == null) return "";

                object[] args = new object[] { propName, "" };
                object result = method.Invoke(pf, args);
                int ret = Convert.ToInt32(result);
                return ret == 0 ? ((args[1] as string) ?? "") : "";
            }
            catch { return ""; }
        }

        private static double RoundDown2(double value)
        {
            return Math.Floor(value * 100.0) / 100.0;
        }

        // ── Lấy tiết diện Pier từ API ─────────────────────────────────────────────
        private SectionInfo FetchPierSection(string pierName, string story)
        {
            int nAreas = 0;
            string[] areaNames = new string[0];
            if (_model.AreaObj.GetNameList(ref nAreas, ref areaNames) != 0 || nAreas == 0)
                return null;

            EnsureStoryElevCache();

            double totalAc = 0;
            double maxLength = 0;
            double thickness = 0;
            string matProp = "";

            foreach (string area in areaNames)
            {
                string p = "";
                if (_model.AreaObj.GetPier(area, ref p) != 0) continue;
                if (!string.Equals(p == null ? "" : p.Trim(), pierName, StringComparison.OrdinalIgnoreCase)) continue;
                if (!AreaBelongsToStory(area, story)) continue;

                double ac, thk, len; string mat;
                if (!TryGetWallAreaInfo(area, out ac, out thk, out len, out mat)) continue;

                totalAc += ac;
                if (len > maxLength) maxLength = len;
                if (thickness <= 0) thickness = thk;
                if (string.IsNullOrWhiteSpace(matProp)) matProp = mat;
            }

            if (totalAc <= 0 || thickness <= 0 || string.IsNullOrWhiteSpace(matProp))
                return null;

            return new SectionInfo
            {
                Ac = maxLength * thickness,
                T3 = RoundDown2(thickness),
                T2 = RoundDown2(maxLength),
                Material = matProp
            };
        }

        // ── Kiểm tra area thuộc tầng ─────────────────────────────────────────────
        private bool AreaBelongsToStory(string areaName, string story)
        {
            EnsureStoryElevCache();

            double zTop, zBot;
            if (_storyElevCache == null
                || !_storyElevCache.TryGetValue(NormalizeStory(story) + "_TOP", out zTop)
                || !_storyElevCache.TryGetValue(NormalizeStory(story) + "_BOT", out zBot))
                return false;

            int nPts = 0;
            string[] pts = new string[0];
            if (_model.AreaObj.GetPoints(areaName, ref nPts, ref pts) != 0 || nPts == 0) return false;

            double zSum = 0;
            foreach (string pt in pts)
            {
                double x = 0, y = 0, z = 0;
                if (_model.PointObj.GetCoordCartesian(pt, ref x, ref y, ref z) != 0) return false;
                zSum += z;
            }
            double zMid = zSum / nPts;
            return zMid >= zBot - 1e-6 && zMid <= zTop + 1e-6;
        }

        // ── Cache cao độ tầng (gọi GetStories một lần) ───────────────────────────
        private void EnsureStoryElevCache()
        {
            if (_storyElevCache != null) return;
            _storyElevCache = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);

            int n = 0;
            string[] sNames = new string[0];
            double[] sElevs = new double[0];
            double[] sHts = new double[0];
            bool[] isMas = new bool[0];
            string[] simTo = new string[0];
            bool[] splice = new bool[0];
            double[] splHt = new double[0];

            if (_model.Story.GetStories(ref n, ref sNames, ref sElevs, ref sHts,
                    ref isMas, ref simTo, ref splice, ref splHt) != 0 || n == 0)
                return;

            for (int i = 0; i < n; i++)
            {
                if (string.IsNullOrWhiteSpace(sNames[i])) continue;
                string key = NormalizeStory(sNames[i]);
                _storyElevCache[key + "_TOP"] = sElevs[i];
                _storyElevCache[key + "_BOT"] = sElevs[i] - sHts[i];
            }
        }

        // ── Lấy thông tin diện tích/bề dày area vách ─────────────────────────────
        private bool TryGetWallAreaInfo(string areaName,
            out double ac, out double thickness, out double length, out string matProp)
        {
            ac = thickness = length = 0; matProp = "";

            string propName = "";
            if (_model.AreaObj.GetProperty(areaName, ref propName) != 0
                || string.IsNullOrWhiteSpace(propName)) return false;

            if (!TryGetWallThickness(propName, out thickness, out matProp)) return false;

            int nPts = 0;
            string[] pts = new string[0];
            if (_model.AreaObj.GetPoints(areaName, ref nPts, ref pts) != 0 || nPts < 2) return false;

            var xs = new List<double>(nPts);
            var ys = new List<double>(nPts);
            foreach (string pt in pts)
            {
                double x = 0, y = 0, z = 0;
                _model.PointObj.GetCoordCartesian(pt, ref x, ref y, ref z);
                xs.Add(x); ys.Add(y);
            }

            double maxDist = 0;
            for (int i = 0; i < xs.Count; i++)
                for (int j = i + 1; j < xs.Count; j++)
                {
                    double dx = xs[i] - xs[j];
                    double dy = ys[i] - ys[j];
                    double d = Math.Sqrt(dx * dx + dy * dy);
                    if (d > maxDist) maxDist = d;
                }

            length = maxDist;
            if (length <= 0 || thickness <= 0) return false;

            ac = thickness * length;
            return true;
        }

        private bool TryGetWallThickness(string propName, out double thickness, out string matProp)
        {
            thickness = 0; matProp = "";

            eWallPropType wType = eWallPropType.Specified;
            eShellType sType = eShellType.ShellThin;
            int c = 0;
            string n = "", g = "";

            int ret = _model.PropArea.GetWall(propName, ref wType, ref sType,
                ref matProp, ref thickness, ref c, ref n, ref g);

            return ret == 0 && thickness > 0 && !string.IsNullOrWhiteSpace(matProp);
        }

        private void ApplyConcreteToSection(SectionInfo sec)
        {
            sec.FckCube = _fckCube;
            sec.Fck = _fck;
            sec.Fcd = _fcd;
        }

        // ── Sắp xếp theo cao độ tầng ETABS ──────────────────────────────────────
        private List<AxialCheckRow> SortByStory(List<AxialCheckRow> rows)
        {
            EnsureStoryElevCache();

            var sorted = rows
                .OrderByDescending(x => GetStoryTop(x.Story))
                .ThenBy(x => GetElementTypeSortOrder(x.ElementType))
                .ThenBy(x => x.Element)
                .ThenBy(x => x.Combo)
                .ToList();

            for (int i = 0; i < sorted.Count; i++)
                sorted[i].STT = i + 1;

            return sorted;
        }

        private double GetStoryTop(string story)
        {
            double e;
            if (_storyElevCache != null && _storyElevCache.TryGetValue(NormalizeStory(story) + "_TOP", out e))
                return e;
            return double.MinValue;
        }

        private static int GetElementTypeSortOrder(string elementType)
        {
            if (string.Equals(elementType, "Column", StringComparison.OrdinalIgnoreCase)) return 0;
            if (string.Equals(elementType, "Pier.Col", StringComparison.OrdinalIgnoreCase)) return 1;
            if (string.Equals(elementType, "Pier.Wall", StringComparison.OrdinalIgnoreCase)) return 2;
            return 3;
        }

        // ── Chọn combo output ─────────────────────────────────────────────────────
        private void SelectOnlyCombo(string combo)
        {
            _model.Results.Setup.DeselectAllCasesAndCombosForOutput();
            _model.Results.Setup.SetComboSelectedForOutput(combo);
        }

        // ── Tạo row kết quả ───────────────────────────────────────────────────────
        private static AxialCheckRow MakeRow(
            ref int stt, string story, string element, string type,
            string combo, SectionInfo sec,
            double ned, double acFcd, double nud, double limit)
        {
            return new AxialCheckRow
            {
                STT = stt++,
                Story = story,
                Element = element,
                ElementType = type,
                Combo = combo,
                Material = sec.Material,
                FckCube = sec.FckCube,
                Fck = sec.Fck,
                Fcd = sec.Fcd,
                Ned = Math.Abs(ned),
                T3 = sec.T3,
                T2 = sec.T2,
                Ac = sec.Ac,
                AcFcd = acFcd,
                NuD = nud,
                VdLimit = limit,
                Result = nud <= limit ? "Thỏa mãn" : "Không thỏa mãn"
            };
        }

        // ── Utilities ──────────────────────────────────────────────────────────────
        private static string NormalizePier(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            s = s.Trim();
            int idx = s.IndexOfAny(new[] { '-', '@', ' ', '\t', ':' });
            if (idx > 0) s = s.Substring(0, idx);
            return s.Trim().ToUpperInvariant();
        }

        private static string NormalizeStory(string s)
        {
            return (s ?? "").Trim().ToUpperInvariant();
        }
    }
}
