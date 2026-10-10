using ETABSv1;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Etabs_Ultimate_Tools
{
    public static class TopDisplacementExtractor
    {
        public static List<TopDisplacementRow> Calculate(
            cSapModel sap, string comboX, string comboY, double limitDenominator)
            => Calculate(sap, comboX, comboY, limitDenominator, DisplacementSource.Diaphragm);

        public static List<TopDisplacementRow> Calculate(
            cSapModel sap, string comboX, string comboY, double limitDenominator,
            DisplacementSource source, string baseStoryOverride = null,
            DiaphragmAggregationMode aggregationMode = DiaphragmAggregationMode.Max)
        {
            var rows = new List<TopDisplacementRow>();
            var stories = EtabsHelper.ReadStories(sap);
            if (stories.Count == 0) return rows;

            var xInfo = string.IsNullOrWhiteSpace(comboX)
                ? Tuple.Create(new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase), new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase))
                : EtabsTableReader.ReadStoryDisplacementInfo(sap, comboX, "X", source, aggregationMode);
            var yInfo = string.IsNullOrWhiteSpace(comboY)
                ? Tuple.Create(new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase), new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase))
                : EtabsTableReader.ReadStoryDisplacementInfo(sap, comboY, "Y", source, aggregationMode);

            var xMap = xInfo.Item1;
            var yMap = yInfo.Item1;
            var xNames = xInfo.Item2;
            var yNames = yInfo.Item2;

            AddStoriesFromDisplacementResults(sap, stories, xMap);
            AddStoriesFromDisplacementResults(sap, stories, yMap);

            double baseElevation = FindFoundationElevation(stories, xMap, yMap, baseStoryOverride);

            var checkStories = stories
                .Select(s => new { Story = s, H = Math.Abs(s.Elevation - baseElevation) })
                .Where(x => x.H > 1e-9 && !EtabsHelper.IsBaseLevel(x.Story.Name))
                .OrderByDescending(x => x.Story.Elevation)
                .ToList();

            foreach (var combo in new[] { (Name: comboX, Dir: "X"), (Name: comboY, Dir: "Y") })
            {
                if (string.IsNullOrWhiteSpace(combo.Name)) continue;
                foreach (var item in checkStories)
                {
                    var chosenMap = combo.Dir.Equals("X", StringComparison.OrdinalIgnoreCase) ? xMap : yMap;
                    var chosenNames = combo.Dir.Equals("X", StringComparison.OrdinalIgnoreCase) ? xNames : yNames;
                    rows.Add(CalculateOne(combo.Name, combo.Dir,
                        item.Story.Name, item.Story.Elevation, item.H, limitDenominator,
                        chosenMap, chosenNames.TryGetValue(item.Story.Name, out var name) ? name : null));
                }
            }
            return rows;
        }

        private static void AddStoriesFromDisplacementResults(
            cSapModel sap, List<EtabsHelper.StoryInfo> stories,
            Dictionary<string, double> resultMap)
        {
            var knownStories = new HashSet<string>(
                stories.Select(s => s.Name), StringComparer.OrdinalIgnoreCase);
            foreach (string storyName in resultMap.Keys)
            {
                if (knownStories.Contains(storyName)) continue;

                double elevation = 0.0, height = 0.0;
                if (sap.Story.GetElevation(storyName, ref elevation) != 0) continue;
                sap.Story.GetHeight(storyName, ref height);

                stories.Add(new EtabsHelper.StoryInfo
                {
                    Name = storyName,
                    Elevation = elevation,
                    Height = height
                });
                knownStories.Add(storyName);
            }
        }

        private static TopDisplacementRow CalculateOne(
            string combo, string dir,
            string topStory, double storyElevation, double h, double limitDenominator,
            Dictionary<string, double> displacements, string sourceTableName)
        {
            double u = displacements.TryGetValue(topStory, out var value) ? value : 0.0;

            double ratio = h > 1e-9 ? Math.Abs(u) / h : 0.0;
            double limit = limitDenominator > 0 ? 1.0 / limitDenominator : 0.0;

            return new TopDisplacementRow
            {
                Direction = dir,
                Combo = combo,
                TopStory = topStory,
                StoryElevation = storyElevation,
                TopElevation = h,
                SourceTableName = sourceTableName,
                TopDisplacement = Math.Abs(u),
                Ratio = ratio,
                LimitDenominator = limitDenominator,
                Check = limit > 0 && ratio > limit ? "NG" : "OK"
            };
        }

        private static double FindFoundationElevation(
            List<EtabsHelper.StoryInfo> stories,
            Dictionary<string, double> xMap, Dictionary<string, double> yMap,
            string baseStoryOverride = null)
        {
            if (stories == null || stories.Count == 0)
                return 0.0;

            if (!string.IsNullOrWhiteSpace(baseStoryOverride))
            {
                string name = baseStoryOverride.Trim();
                var exact = stories.FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));
                if (exact != null) return exact.Elevation;

                var contains = stories.FirstOrDefault(s => s.Name.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0);
                if (contains != null) return contains.Elevation;
            }

            var foundationLike = stories
                .Where(s => s.Name.IndexOf("Base", StringComparison.OrdinalIgnoreCase) >= 0
                    || s.Name.IndexOf("FDT", StringComparison.OrdinalIgnoreCase) >= 0
                    || s.Name.IndexOf("Foundation", StringComparison.OrdinalIgnoreCase) >= 0
                    || s.Name.IndexOf("Móng", StringComparison.OrdinalIgnoreCase) >= 0
                    || s.Name.IndexOf("Footing", StringComparison.OrdinalIgnoreCase) >= 0
                    || s.Name.IndexOf("Pilecap", StringComparison.OrdinalIgnoreCase) >= 0)
                .OrderBy(s => s.Elevation)
                .ToList();
            if (foundationLike.Count > 0)
                return foundationLike.First().Elevation;

            const double zeroTol = 1e-9;
            foreach (var st in stories.OrderBy(s => s.Elevation))
            {
                bool hasData = false;
                bool isFixed = true;

                if (xMap.TryGetValue(st.Name, out var ux))
                {
                    hasData = true;
                    if (Math.Abs(ux) > zeroTol) isFixed = false;
                }

                if (yMap.TryGetValue(st.Name, out var uy))
                {
                    hasData = true;
                    if (Math.Abs(uy) > zeroTol) isFixed = false;
                }

                if (hasData && isFixed)
                    return st.Elevation;
            }

            foreach (var st in stories.OrderBy(s => s.Elevation))
                if (xMap.ContainsKey(st.Name) || yMap.ContainsKey(st.Name))
                    return st.Elevation;

            return stories.Min(s => s.Elevation);
        }
    }
}
