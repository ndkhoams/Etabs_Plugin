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
        {
            var rows = new List<SeismicDriftRow>();
            var stories = EtabsHelper.ReadStories(sap);
            if (stories.Count == 0) return rows;

            if (q <= 0) q = 1.0;
            if (nu <= 0) nu = 1.0;

            // Ngưỡng drift cho phép: drift ≤ limit/(ν·q)
            double allow = (q * nu) > 0 ? limitRatio / (q * nu) : 0.0;

            foreach (var combo in new[] { (Name: comboX, Dir: "X"), (Name: comboY, Dir: "Y") })
            {
                if (string.IsNullOrWhiteSpace(combo.Name)) continue;

                var drifts = ReadStoryDrifts(sap, combo.Name, combo.Dir);

                foreach (var st in stories.OrderByDescending(s => s.Elevation))
                {
                    if (EtabsHelper.IsBaseLevel(st.Name)) continue;
                    if (st.Height <= 0) continue;

                    double drift = drifts.TryGetValue(st.Name, out var d) ? d : 0.0;

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
                        Check = allow > 0 && drift > allow ? "NG" : "OK"
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
