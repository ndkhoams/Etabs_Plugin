using ETABSv1;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Etabs_Ultimate_Tools
{
    public static class PDeltaExtractor
    {
        public static List<PDeltaCheckRow> Calculate(
            cSapModel sap, string driftCombo, string vCombo, string dir, double q)
        {
            if (string.IsNullOrWhiteSpace(driftCombo) || string.IsNullOrWhiteSpace(vCombo))
                return new List<PDeltaCheckRow>();

            var stories = EtabsHelper.ReadStories(sap);
            var drifts = ReadStoryDrifts(sap, driftCombo, dir);
            var shears = ReadStoryForces(sap, vCombo, dir);
            var pTot = ReadPtotFromMassSummary(sap, stories);

            var rows = new List<PDeltaCheckRow>();
            foreach (var st in stories.OrderByDescending(x => x.Elevation))
            {
                if (st.Height <= 0) continue;

                double driftElastic = GetOrZero(drifts, st.Name);
                double dr = q * driftElastic;
                double vtot = GetOrZero(shears, st.Name);

                if (vtot <= 1.0) continue;

                double ptot = GetOrZero(pTot, st.Name);
                double theta = ptot * dr / vtot;

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

        public static List<string> GetLoadCombinations(cSapModel sap)
            => EtabsHelper.GetLoadCombinations(sap);

        private static string GetConclusion(double theta)
        {
            if (theta <= 0.10) return "OK - bỏ qua hiệu ứng bậc 2";
            if (theta <= 0.20) return "OK - nhân nội lực với 1/(1-θ)";
            if (theta <= 0.30) return "Cần xét P-Delta chính xác / kiểm soát";
            return "NG - θ > 0.30, cần tăng độ cứng/thiết kế lại";
        }

        private static Dictionary<string, double> ReadPtotFromMassSummary(
            cSapModel sap, List<EtabsHelper.StoryInfo> stories)
        {
            var table = EtabsTableReader.ReadTableWithFallback(
                sap, EtabsTableReader.MassSummaryTableNames, "");

            var storyMass = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            foreach (var row in table)
            {
                string story = EtabsTableReader.Get(row, "Story", "StoryName", "Story Name");
                if (string.IsNullOrWhiteSpace(story)) continue;

                double mx = Math.Abs(EtabsTableReader.GetDouble(row,
                    "Mass X", "MassX", "MassUX", "UX", "U1", "X Mass", "Mass in X", "Mass X kN-s²/m"));
                double my = Math.Abs(EtabsTableReader.GetDouble(row,
                    "Mass Y", "MassY", "MassUY", "UY", "U2", "Y Mass", "Mass in Y", "Mass Y kN-s²/m"));

                double m = Math.Max(mx, my);
                if (m <= 0)
                    m = Math.Abs(EtabsTableReader.GetDouble(row,
                        "Mass", "Total Mass", "Story Mass", "Mass kN-s²/m"));
                if (m <= 0) continue;

                if (storyMass.TryGetValue(story, out var cur)) storyMass[story] = cur + m;
                else storyMass[story] = m;
            }

            const double g = 9.80665;
            var storyWeight = storyMass.ToDictionary(
                kv => kv.Key, kv => kv.Value * g, StringComparer.OrdinalIgnoreCase);

            var result = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            double cumulative = 0.0;
            foreach (var st in stories.OrderByDescending(s => s.Elevation))
            {
                cumulative += GetOrZero(storyWeight, st.Name);
                result[st.Name] = cumulative;
            }
            return result;
        }

        private static Dictionary<string, double> ReadStoryDrifts(
            cSapModel sap, string loadCase, string dir)
        {
            EtabsHelper.SelectCaseOrCombo(sap, loadCase);

            int n = 0;
            string[] story = null, caseName = null, stepType = null, dirArr = null, label = null;
            double[] stepNum = null, drift = null, x = null, y = null, z = null;
            sap.Results.StoryDrifts(ref n, ref story, ref caseName, ref stepType,
                ref stepNum, ref dirArr, ref drift, ref label, ref x, ref y, ref z);

            var result = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < n; i++)
            {
                if (!string.Equals(caseName[i], loadCase, StringComparison.OrdinalIgnoreCase)) continue;
                if (!dirArr[i].StartsWith(dir, StringComparison.OrdinalIgnoreCase)) continue;

                double val = Math.Abs(drift[i]);
                if (!result.TryGetValue(story[i], out var cur) || val > cur)
                    result[story[i]] = val;
            }
            return result;
        }

        private static Dictionary<string, double> ReadStoryForces(
            cSapModel sap, string loadCase, string dir)
        {
            string forceField = dir.Equals("X", StringComparison.OrdinalIgnoreCase) ? "VX" : "VY";

            EtabsHelper.SelectComboOnly(sap, loadCase);
            var table = EtabsTableReader.ReadTable(sap, "Story Forces", loadCase);

            var result = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            foreach (var row in table)
            {
                string story = EtabsTableReader.Get(row, "Story", "StoryName");
                if (string.IsNullOrWhiteSpace(story)) continue;

                string outputCase = EtabsTableReader.Get(row,
                    "Output Case", "OutputCase", "Load Case", "LoadCase", "Case", "Combo", "Combination");
                if (!EtabsHelper.IsSameOrBlank(outputCase, loadCase)) continue;

                if (!IsBottom(EtabsTableReader.Get(row, "Location"))) continue;

                double v = Math.Abs(EtabsTableReader.GetDouble(row,
                    forceField, "V" + dir, forceField + " kN", "V" + dir + " kN", "F" + dir));

                if (!result.TryGetValue(story, out var cur) || v > cur)
                    result[story] = v;
            }
            return result;
        }

        private static bool IsBottom(string location)
        {
            if (string.IsNullOrWhiteSpace(location)) return true;
            string s = location.Trim();
            return s.Equals("Bot", StringComparison.OrdinalIgnoreCase)
                || s.Equals("B", StringComparison.OrdinalIgnoreCase)
                || s.IndexOf("Bottom", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static double GetOrZero(Dictionary<string, double> dict, string key)
            => dict.TryGetValue(key, out var v) ? v : 0.0;
    }
}
