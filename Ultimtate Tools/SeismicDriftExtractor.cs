using ETABSv1;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Etabs_Ultimate_Tools
{
    /// <summary>
    /// Kiểm tra chuyển vị lệch tầng do tải trọng động đất theo TCVN 9386-1:2025.
    /// Điều kiện hạn chế hư hỏng: dr · ν ≤ limit · h  ⇔  drift ≤ limit/(ν·q)
    /// với drift = de/h lấy trực tiếp từ ETABS Story Drifts (tổ hợp động đất thuần 1.0EX + 0.3EY, đàn hồi),
    /// lấy max trị tuyệt đối theo từng tầng & từng phương.
    /// </summary>
    public static class SeismicDriftExtractor
    {
        public static List<SeismicDriftRow> Calculate(
            cSapModel sap, string comboX, string comboY, double q, double nu, double limitRatio)
            => Calculate(sap, comboX, comboY, q, nu, limitRatio, null);

        public static List<SeismicDriftRow> Calculate(
            cSapModel sap, string comboX, string comboY, double q, double nu, double limitRatio,
            DisplacementSource? source,
            DiaphragmAggregationMode aggregationMode = DiaphragmAggregationMode.Max)
        {
            var rows = new List<SeismicDriftRow>();
            var stories = EtabsHelper.ReadStories(sap);
            if (stories.Count == 0) return rows;
            EtabsHelper.AddMissingStoriesFromDrifts(sap, stories, comboX, comboY);

            if (q <= 0) q = 1.0;
            if (nu <= 0) nu = 1.0;

            // Ngưỡng drift cho phép: drift ≤ limit/(ν·q)
            double allow = (q * nu) > 0 ? limitRatio / (q * nu) : 0.0;

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

                    bool hasDrift = drifts.TryGetValue(st.Name, out double drift);

                    rows.Add(new SeismicDriftRow
                    {
                        Direction = combo.Dir,
                        Combo = combo.Name,
                        Story = st.Name,
                        Elevation = st.Elevation,
                        Height = st.Height,
                        Q = q,
                        Nu = nu,
                        Drift = drift,
                        LimitRatio = limitRatio,
                        Check = !hasDrift || allow <= 0 ? "Không có dữ liệu" : (drift > allow ? "NG" : "OK")
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
            int ret = sap.Results.StoryDrifts(ref n, ref story, ref caseName, ref stepType,
                ref stepNum, ref dirArr, ref drift, ref label, ref x, ref y, ref z);

            var result = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            // API lỗi hoặc trả mảng thiếu -> coi như không có dữ liệu (nơi gọi sẽ đánh dấu
            // "Không có dữ liệu") thay vì để tràn index / NullReference.
            if (ret != 0 || story == null || caseName == null || dirArr == null || drift == null)
                return result;

            int count = Math.Min(n, Math.Min(story.Length,
                Math.Min(caseName.Length, Math.Min(dirArr.Length, drift.Length))));
            for (int i = 0; i < count; i++)
            {
                if (story[i] == null || caseName[i] == null || dirArr[i] == null) continue;
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
