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
                    : EtabsHelper.ReadStoryDrifts(sap, combo.Name, combo.Dir);

                foreach (var st in stories.OrderByDescending(s => s.Elevation))
                {
                    if (EtabsHelper.IsBaseLevel(st.Name)) continue;
                    if (st.Height <= 0) continue;

                    bool hasDrift = drifts.TryGetValue(st.Name, out double drift);

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
                        Check = !hasDrift || limit <= 0 ? "Không có dữ liệu" : (drift > limit ? "NG" : "OK")
                    });
                }
            }
            return rows;
        }

    }
}
