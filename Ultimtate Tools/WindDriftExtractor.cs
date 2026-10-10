using ETABSv1;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Etabs_Ultimate_Tools
{
    /// <summary>
    /// Kiểm tra chuyển vị lệch tầng (inter-story drift) do tải trọng gió.
    /// Dùng trực tiếp Story Drifts của ETABS (đã là tỉ số Δ/h theo từng tầng),
    /// lấy max trị tuyệt đối theo từng tầng & từng phương.
    /// </summary>
    public static class WindDriftExtractor
    {
        public static List<WindDriftRow> Calculate(
            cSapModel sap, string comboX, string comboY, double limitDenominator)
            => Calculate(sap, comboX, comboY, limitDenominator, null);

        public static List<WindDriftRow> Calculate(
            cSapModel sap, string comboX, string comboY, double limitDenominator,
            DisplacementSource? source,
            DiaphragmAggregationMode aggregationMode = DiaphragmAggregationMode.Max)
        {
            var rows = new List<WindDriftRow>();
            var stories = EtabsHelper.ReadStories(sap);
            if (stories.Count == 0) return rows;
            EtabsHelper.AddMissingStoriesFromDrifts(sap, stories, comboX, comboY);

            double limit = limitDenominator > 0 ? 1.0 / limitDenominator : 0.0;

            foreach (var combo in new[] { (Name: comboX, Dir: "X"), (Name: comboY, Dir: "Y") })
            {
                if (string.IsNullOrWhiteSpace(combo.Name)) continue;

                var drifts = source.HasValue
                    ? StoryDisplacementCalculator.ReadDrifts(sap, combo.Name, combo.Dir, stories, source.Value, aggregationMode)
                    : ReadStoryDrifts(sap, combo.Name, combo.Dir);

                foreach (var st in stories.OrderByDescending(s => s.Elevation))
                {
                    if (EtabsHelper.IsBaseLevel(st.Name)) continue;
                    if (st.Height <= 0) continue;

                    double drift = drifts.TryGetValue(st.Name, out var d) ? d : 0.0;

                    rows.Add(new WindDriftRow
                    {
                        Direction = combo.Dir,
                        Combo = combo.Name,
                        Story = st.Name,
                        Elevation = st.Elevation,
                        Height = st.Height,
                        Drift = drift,
                        DriftDisplacement = drift * st.Height,
                        LimitDenominator = limitDenominator,
                        Check = limit > 0 && drift > limit ? "NG" : "OK"
                    });
                }
            }
            return rows;
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
    }
}
