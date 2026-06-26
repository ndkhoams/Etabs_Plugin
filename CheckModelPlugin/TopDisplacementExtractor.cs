using ETABSv1;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Etabs_Ultimate_Tools
{
    public static class TopDisplacementExtractor
    {
        private const double MmGuessThresholdMeters = 5.0;

        public static List<TopDisplacementRow> Calculate(
            cSapModel sap, string comboX, string comboY, double limitDenominator)
        {
            var rows = new List<TopDisplacementRow>();
            var stories = EtabsHelper.ReadStories(sap);
            if (stories.Count == 0) return rows;

            double baseElevation = FindFoundationElevation(sap, stories, comboX, comboY);

            var checkStories = stories
                .Select(s => new { Story = s, H = Math.Abs(s.Elevation - baseElevation) })
                .Where(x => x.H > 1e-9 && !EtabsHelper.IsBaseLevel(x.Story.Name))
                .OrderByDescending(x => x.Story.Elevation)
                .ToList();

            foreach (var combo in new[] { (Name: comboX, Dir: "X"), (Name: comboY, Dir: "Y") })
            {
                if (string.IsNullOrWhiteSpace(combo.Name)) continue;
                foreach (var item in checkStories)
                    rows.Add(CalculateOne(sap, combo.Name, combo.Dir,
                        item.Story.Name, item.Story.Elevation, item.H, limitDenominator));
            }
            return rows;
        }

        private static TopDisplacementRow CalculateOne(
            cSapModel sap, string combo, string dir,
            string topStory, double storyElevation, double h, double limitDenominator)
        {
            double u = ReadDisplacement(sap, combo, dir, topStory,
                EtabsTableReader.DiaphragmDisplacementTableNames);

            if (Math.Abs(u) < 1e-12)
                u = ReadDisplacement(sap, combo, dir, topStory,
                    EtabsTableReader.StoryDisplacementTableNames);

            double ratio = h > 1e-9 ? Math.Abs(u) / h : 0.0;
            double limit = limitDenominator > 0 ? 1.0 / limitDenominator : 0.0;

            return new TopDisplacementRow
            {
                Direction = dir,
                Combo = combo,
                TopStory = topStory,
                StoryElevation = storyElevation,
                TopElevation = h,
                TopDisplacement = Math.Abs(u),
                Ratio = ratio,
                LimitDenominator = limitDenominator,
                Check = limit > 0 && ratio > limit ? "NG" : "OK"
            };
        }

        private static double FindFoundationElevation(
            cSapModel sap, List<EtabsHelper.StoryInfo> stories,
            string comboX, string comboY)
        {
            const double zeroTol = 1e-9;

            var xMap = string.IsNullOrWhiteSpace(comboX)
                ? new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase)
                : ReadDisplacementMap(sap, comboX, "X", EtabsTableReader.DiaphragmDisplacementTableNames);

            var yMap = string.IsNullOrWhiteSpace(comboY)
                ? new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase)
                : ReadDisplacementMap(sap, comboY, "Y", EtabsTableReader.DiaphragmDisplacementTableNames);

            foreach (var st in stories.OrderBy(s => s.Elevation))
            {
                bool hasData = false, isFixed = true;
                if (xMap.TryGetValue(st.Name, out var ux)) { hasData = true; if (Math.Abs(ux) > zeroTol) isFixed = false; }
                if (yMap.TryGetValue(st.Name, out var uy)) { hasData = true; if (Math.Abs(uy) > zeroTol) isFixed = false; }
                if (hasData && isFixed) return st.Elevation;
            }

            foreach (var st in stories.OrderBy(s => s.Elevation))
                if (xMap.ContainsKey(st.Name) || yMap.ContainsKey(st.Name))
                    return st.Elevation;

            return stories.Min(s => s.Elevation);
        }

        private static double ReadDisplacement(
            cSapModel sap, string combo, string dir, string topStory, string[] tableNames)
        {
            double bestWithCase = 0.0, bestAnyCase = 0.0;

            foreach (var tableName in tableNames)
            {
                List<Dictionary<string, string>> table;
                try { table = EtabsTableReader.ReadTable(sap, tableName, combo); }
                catch { continue; }

                foreach (var row in table)
                {
                    string story = EtabsTableReader.Get(row, "Story", "StoryName", "Story Name", "Level");
                    if (!string.IsNullOrWhiteSpace(story) &&
                        !string.Equals(story.Trim(), topStory.Trim(), StringComparison.OrdinalIgnoreCase))
                        continue;

                    string outputCase = EtabsTableReader.Get(row,
                        "Output Case", "OutputCase", "Load Case", "LoadCase", "Case", "Combo", "Combination");
                    double u = NormalizeToMeters(ReadDirectionalDisplacement(row, dir));

                    if (u > bestAnyCase) bestAnyCase = u;
                    if (EtabsHelper.IsSameOrBlank(outputCase, combo) && u > bestWithCase) bestWithCase = u;
                }

                if (bestWithCase > 0) return bestWithCase;
            }
            return bestAnyCase;
        }

        private static Dictionary<string, double> ReadDisplacementMap(
            cSapModel sap, string combo, string dir, string[] tableNames)
        {
            var bestWithCase = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            var bestAnyCase = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);

            foreach (var tableName in tableNames)
            {
                List<Dictionary<string, string>> table;
                try { table = EtabsTableReader.ReadTable(sap, tableName, combo); }
                catch { continue; }

                foreach (var row in table)
                {
                    string story = EtabsTableReader.Get(row, "Story", "StoryName", "Story Name", "Level");
                    if (string.IsNullOrWhiteSpace(story)) continue;
                    story = story.Trim();

                    string outputCase = EtabsTableReader.Get(row,
                        "Output Case", "OutputCase", "Load Case", "LoadCase", "Case", "Combo", "Combination");
                    double u = NormalizeToMeters(ReadDirectionalDisplacement(row, dir));

                    if (!bestAnyCase.TryGetValue(story, out var anyCur) || u > anyCur) bestAnyCase[story] = u;
                    if (EtabsHelper.IsSameOrBlank(outputCase, combo))
                        if (!bestWithCase.TryGetValue(story, out var wcCur) || u > wcCur) bestWithCase[story] = u;
                }

                if (bestWithCase.Count > 0) return bestWithCase;
            }
            return bestAnyCase;
        }

        private static double ReadDirectionalDisplacement(Dictionary<string, string> row, string dir)
        {
            return dir.Equals("X", StringComparison.OrdinalIgnoreCase)
                ? EtabsTableReader.GetDouble(row,
                    "UX", "Ux", "UX m", "Ux m", "UX mm", "Ux mm",
                    "U1", "U1 m", "U1 mm",
                    "X", "X m", "X mm",
                    "X-Displ", "X Displ", "Displ X", "Translation X",
                    "Global X", "GlobalX", "X Translation")
                : EtabsTableReader.GetDouble(row,
                    "UY", "Uy", "UY m", "Uy m", "UY mm", "Uy mm",
                    "U2", "U2 m", "U2 mm",
                    "Y", "Y m", "Y mm",
                    "Y-Displ", "Y Displ", "Displ Y", "Translation Y",
                    "Global Y", "GlobalY", "Y Translation");
        }

        /// <summary>Đổi giá trị từ mm về m nếu cần, và lấy trị tuyệt đối.</summary>
        private static double NormalizeToMeters(double u)
        {
            u = Math.Abs(u);
            if (u > MmGuessThresholdMeters) u /= 1000.0;
            return u;
        }
    }
}
