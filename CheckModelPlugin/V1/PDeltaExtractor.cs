using ETABSv1;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace CheckModelPlugin
{
    public static class PDeltaExtractor
    {
        public static List<PDeltaCheckRow> Calculate(cSapModel sap, string driftCombo, string vCombo, string dir, double q)
        {
            if (string.IsNullOrWhiteSpace(driftCombo)) return new List<PDeltaCheckRow>();
            if (string.IsNullOrWhiteSpace(vCombo)) return new List<PDeltaCheckRow>();

            var stories = ReadStories(sap);
            var drifts = ReadStoryDrifts(sap, driftCombo, dir);
            var shears = ReadStoryForces(sap, vCombo, dir);

            // Ptot lấy từ Mass Summary by Story, đổi Mass × g rồi cộng dồn từ mái xuống.
            // Giữ nguyên cách đọc/lọc drift và Vtot của file gốc.
            var pTot = ReadPtotFromMassSummary(sap, stories);

            var rows = new List<PDeltaCheckRow>();
            foreach (var st in stories.OrderByDescending(x => x.Elevation))
            {
                if (st.Height <= 0) continue;

                double driftElastic = Math.Abs(GetOrZero(drifts, st.Name));
                double dr = q * driftElastic;
                double vtot = Math.Abs(GetOrZero(shears, st.Name));

                // Bỏ qua tầng/ngàm đáy khi ETABS trả Vtot xấp xỉ 0.
                // Trong ETABS, giá trị rất nhỏ kiểu 1E-6 vẫn có thể hiển thị là 0 trên form,
                // nhưng nếu đem chia sẽ tạo θ rất lớn/∞ không có ý nghĩa thiết kế.
                // Dùng ngưỡng 1 kN để bỏ qua các tầng có lực cắt đáy gần bằng 0.
                if (vtot <= 1.0)
                    continue;

                double ptot = Math.Abs(GetOrZero(pTot, st.Name));
                // ETABS StoryDrifts trả về tỷ số lệch tầng dre = Δ/h (không thứ nguyên), không phải chuyển vị Δ (m).
                // Do đó dr = q × dre cũng là tỷ số lệch tầng thiết kế.
                // Công thức EC8 θ = Ptot × Δd / (Vtot × h) được viết lại thành:
                // θ = Ptot × (Δd/h) / Vtot = Ptot × dr / Vtot.
                // Không chia thêm cho h để tránh chia h hai lần.
                double theta = (vtot > 1e-9) ? ptot * dr / vtot : 0;

                rows.Add(new PDeltaCheckRow
                {
                    Direction = dir,
                    Story = st.Name,
                    Elevation = st.Elevation,
                    Height = st.Height,
                    ElasticDrift = driftElastic,
                    DesignDrift = dr,
                    Ptot = ptot,
                    Vtot = vtot,
                    Theta = theta,
                    Amplification = theta < 1.0 ? 1.0 / (1.0 - theta) : double.PositiveInfinity,
                    Conclusion = GetConclusion(theta)
                });
            }

            return rows;
        }

        private static string GetConclusion(double theta)
        {
            if (theta <= 0.10) return "OK - bỏ qua hiệu ứng bậc 2";
            if (theta <= 0.20) return "OK - nhân nội lực với 1/(1-θ)";
            if (theta <= 0.30) return "Cần xét P-Delta chính xác / kiểm soát";
            return "NG - θ > 0.30, cần tăng độ cứng/thiết kế lại";
        }

        private static double GetOrZero(Dictionary<string, double> dict, string key)
        {
            return dict.TryGetValue(key, out var v) ? v : 0.0;
        }

        private class StoryInfo
        {
            public string Name;
            public double Elevation;
            public double Height;
        }

        private class StoryForceRecord
        {
            public string Story;
            public string OutputCase;
            public string Location;
            public double P;
            public double VX;
            public double VY;
        }

        private static List<StoryInfo> ReadStories(cSapModel sap)
        {
            int numberStories = 0;
            string[] storyNames = null;
            double[] storyElevations = null;
            double[] storyHeights = null;
            bool[] isMasterStory = null;
            string[] similarToStory = null;
            bool[] spliceAbove = null;
            double[] spliceHeight = null;
            sap.Story.GetStories(ref numberStories, ref storyNames, ref storyElevations, ref storyHeights,
                ref isMasterStory, ref similarToStory, ref spliceAbove, ref spliceHeight);

            var list = new List<StoryInfo>();
            for (int i = 0; i < numberStories; i++)
            {
                list.Add(new StoryInfo
                {
                    Name = storyNames[i],
                    Elevation = storyElevations[i],
                    Height = storyHeights[i]
                });
            }
            return list.OrderBy(x => x.Elevation).ToList();
        }


        private static Dictionary<string, double> ReadPtotFromMassSummary(cSapModel sap, List<StoryInfo> stories)
        {
            var storyMass = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);

            // Bảng này không phụ thuộc load combo. ETABS các bản có thể dùng tên field hơi khác nhau,
            // nên đọc mềm nhiều tên cột phổ biến.
            var table = EtabsTableReader.ReadTable(sap, "Mass Summary by Story", "");
            if (table.Count == 0) table = EtabsTableReader.ReadTable(sap, "Story Mass Summary", "");
            if (table.Count == 0) table = EtabsTableReader.ReadTable(sap, "Masses by Story", "");
            if (table.Count == 0) table = EtabsTableReader.ReadTable(sap, "Center Of Mass And Rigidity", "");
            if (table.Count == 0) table = EtabsTableReader.ReadTable(sap, "Centers Of Mass And Rigidity", "");

            foreach (var row in table)
            {
                string story = EtabsTableReader.Get(row, "Story", "StoryName", "Story Name");
                if (string.IsNullOrWhiteSpace(story)) continue;

                double mx = Math.Abs(EtabsTableReader.GetDouble(row,
                    "Mass X", "MassX", "MassUX", "UX", "U1", "X Mass", "Mass in X", "Mass X kN-s²/m"));

                double my = Math.Abs(EtabsTableReader.GetDouble(row,
                    "Mass Y", "MassY", "MassUY", "UY", "U2", "Y Mass", "Mass in Y", "Mass Y kN-s²/m"));

                double m = Math.Max(mx, my);

                // Fallback cho trường hợp ETABS chỉ có một cột Mass.
                if (m <= 0)
                {
                    m = Math.Abs(EtabsTableReader.GetDouble(row,
                        "Mass", "Total Mass", "Story Mass", "Mass kN-s²/m"));
                }

                if (m <= 0) continue;

                if (storyMass.ContainsKey(story))
                    storyMass[story] += m;
                else
                    storyMass[story] = m;
            }

            // Đổi mass sang trọng lượng tầng Wi = mi × g.
            const double g = 9.80665;
            var storyWeight = storyMass.ToDictionary(kv => kv.Key, kv => kv.Value * g, StringComparer.OrdinalIgnoreCase);

            // Ptot tại tầng = tổng trọng lượng tầng đang xét và các tầng phía trên.
            var result = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            double cumulative = 0.0;

            foreach (var st in stories.OrderByDescending(s => s.Elevation))
            {
                cumulative += GetOrZero(storyWeight, st.Name);
                result[st.Name] = cumulative;
            }

            return result;
        }

        private static Dictionary<string, double> ReadStoryDrifts(cSapModel sap, string loadCase, string dir)
        {
            sap.Results.Setup.DeselectAllCasesAndCombosForOutput();
            SelectCaseOrCombo(sap, loadCase);

            int n = 0;
            string[] story = null, caseName = null, stepType = null, dirArr = null, label = null;
            double[] stepNum = null, drift = null, x = null, y = null, z = null;

            // ETABS API v1: Results.StoryDrifts. Nếu bản ETABS báo sai tham số, xem F12 và chỉnh đúng chữ ký hàm.
            sap.Results.StoryDrifts(ref n, ref story, ref caseName, ref stepType, ref stepNum, ref dirArr, ref drift, ref label, ref x, ref y, ref z);

            var result = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < n; i++)
            {
                if (!string.Equals(caseName[i], loadCase, StringComparison.OrdinalIgnoreCase)) continue;
                if (!dirArr[i].StartsWith(dir, StringComparison.OrdinalIgnoreCase)) continue;
                double val = Math.Abs(drift[i]);
                if (!result.ContainsKey(story[i]) || val > result[story[i]]) result[story[i]] = val;
            }
            return result;
        }

        private static Dictionary<string, double> ReadStoryForces(cSapModel sap, string loadCase, string dir)
        {
            // Vtot lấy từ bảng Story Forces theo Load Combination người dùng chọn.
            // ETABSv1.dll của một số bản không có Results.StoryForces, nên dùng DatabaseTables để tương thích.
            var result = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            string forceField = dir.Equals("X", StringComparison.OrdinalIgnoreCase) ? "VX" : "VY";

            sap.Results.Setup.DeselectAllCasesAndCombosForOutput();
            SelectComboOnly(sap, loadCase);
            var table = EtabsTableReader.ReadTable(sap, "Story Forces", loadCase);

            foreach (var row in table)
            {
                var story = EtabsTableReader.Get(row, "Story", "StoryName");
                if (string.IsNullOrWhiteSpace(story)) continue;

                var outputCase = EtabsTableReader.Get(row, "Output Case", "OutputCase", "Load Case", "LoadCase", "Case", "Combo", "Combination");
                if (!IsSameOrBlank(outputCase, loadCase)) continue;

                var location = EtabsTableReader.Get(row, "Location");
                if (!IsBottom(location)) continue;

                double v = Math.Abs(EtabsTableReader.GetDouble(row, forceField, "V" + dir, forceField + " kN", "V" + dir + " kN", "F" + dir));
                if (!result.ContainsKey(story) || v > result[story]) result[story] = v;
            }
            return result;
        }

        private static bool IsBottom(string location)
        {
            if (string.IsNullOrWhiteSpace(location)) return true;
            string s = location.Trim();
            return s.Equals("Bottom", StringComparison.OrdinalIgnoreCase) ||
                   s.Equals("Bot", StringComparison.OrdinalIgnoreCase) ||
                   s.Equals("B", StringComparison.OrdinalIgnoreCase) ||
                   s.IndexOf("Bottom", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool IsSameOrBlank(string outputCase, string selectedName)
        {
            if (string.IsNullOrWhiteSpace(outputCase)) return true;
            if (string.Equals(outputCase.Trim(), selectedName.Trim(), StringComparison.OrdinalIgnoreCase)) return true;
            // Với envelope/min/max combo, ETABS đôi khi thêm StepType hoặc hậu tố vào tên output.
            return outputCase.IndexOf(selectedName, StringComparison.OrdinalIgnoreCase) >= 0 ||
                   selectedName.IndexOf(outputCase, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static void SelectCaseOrCombo(cSapModel sap, string name)
        {
            // Dùng cho drift: cho phép chọn Load Case hoặc Load Combination.
            int ret = sap.Results.Setup.SetCaseSelectedForOutput(name);
            if (ret != 0) sap.Results.Setup.SetComboSelectedForOutput(name);
        }

        private static void SelectComboOnly(cSapModel sap, string comboName)
        {
            // Dùng cho Ptot và Vtot: bắt buộc lấy từ Load Combination người dùng đã định nghĩa.
            sap.Results.Setup.DeselectAllCasesAndCombosForOutput();
            sap.Results.Setup.SetComboSelectedForOutput(comboName);
        }

        public static List<string> GetLoadCombinations(cSapModel sap)
        {
            int n = 0;
            string[] names = null;
            sap.RespCombo.GetNameList(ref n, ref names);
            return names == null ? new List<string>() : names.OrderBy(x => x).ToList();
        }
    }
}
