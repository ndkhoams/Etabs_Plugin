using ETABSv1;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Etabs_Ultimate_Tools
{
    internal static class StoryDisplacementCalculator
    {
        public static Dictionary<string, double> ReadDrifts(
            cSapModel sap, string combo, string dir,
            List<EtabsHelper.StoryInfo> stories, DisplacementSource source,
            DiaphragmAggregationMode aggregationMode = DiaphragmAggregationMode.Max)
        {
            if (source == DisplacementSource.Story)
            {
                EtabsHelper.AddMissingStoriesFromDrifts(sap, stories, combo);
                return ReadEtabsStoryDrifts(sap, combo, dir);
            }

            var diaphragmDrifts = EtabsTableReader.ReadDiaphragmDrifts(
                sap, combo, dir, aggregationMode);
            if (diaphragmDrifts.Count > 0)
            {
                AddMissingStories(sap, stories, diaphragmDrifts);
                return diaphragmDrifts;
            }

            var displacements = EtabsTableReader.ReadStoryDisplacements(sap, combo, dir, source, aggregationMode);
            AddMissingStories(sap, stories, displacements);

            if (displacements.Count == 0)
                return ReadEtabsStoryDrifts(sap, combo, dir);

            var orderedStories = stories.OrderBy(s => s.Elevation).ToList();
            var drifts = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            foreach (var story in orderedStories)
            {
                if (EtabsHelper.IsBaseLevel(story.Name) || story.Height <= 0 ||
                    !displacements.TryGetValue(story.Name, out var current))
                    continue;

                var lowerStory = orderedStories
                    .Where(s => s.Elevation < story.Elevation - 1e-9 &&
                                displacements.ContainsKey(s.Name))
                    .OrderByDescending(s => s.Elevation)
                    .FirstOrDefault();
                double lower = lowerStory != null ? displacements[lowerStory.Name] : 0.0;
                drifts[story.Name] = Math.Abs(current - lower) / story.Height;
            }

            if (drifts.Count == 0)
                return ReadEtabsStoryDrifts(sap, combo, dir);

            return drifts;
        }

        private static Dictionary<string, double> ReadEtabsStoryDrifts(
            cSapModel sap, string combo, string dir)
        {
            var table = EtabsTableReader.ReadTableWithFallback(
                sap, EtabsTableReader.StoryDriftTableNames, combo);
            var tableResult = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            foreach (var row in table)
            {
                string story = EtabsTableReader.Get(row, "Story", "StoryName", "Story Name", "Level");
                if (string.IsNullOrWhiteSpace(story)) continue;

                string outputCase = EtabsTableReader.Get(row,
                    "Output Case", "OutputCase", "Load Case", "LoadCase", "Case", "Combo", "Combination");
                if (!EtabsHelper.IsSameOrBlank(outputCase, combo)) continue;

                string direction = EtabsTableReader.Get(row, "Direction", "Dir");
                if (string.IsNullOrWhiteSpace(direction) ||
                    !direction.Trim().StartsWith(dir, StringComparison.OrdinalIgnoreCase))
                    continue;

                double value = Math.Abs(EtabsTableReader.GetDouble(row,
                    "Drift", "Story Drift", "Drift Ratio", "Ratio"));
                if (!tableResult.TryGetValue(story, out var current) || value > current)
                    tableResult[story] = value;
            }

            if (tableResult.Count > 0)
                return tableResult;

            EtabsHelper.SelectCaseOrCombo(sap, combo);

            int count = 0;
            string[] stories = null, cases = null, stepTypes = null;
            string[] directions = null, labels = null;
            double[] stepNumbers = null, drifts = null, x = null, y = null, z = null;
            int ret = sap.Results.StoryDrifts(ref count, ref stories, ref cases,
                ref stepTypes, ref stepNumbers, ref directions, ref drifts,
                ref labels, ref x, ref y, ref z);

            var result = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            if (ret != 0 || stories == null || cases == null || directions == null || drifts == null)
                return result;

            for (int i = 0; i < count && i < stories.Length && i < cases.Length &&
                i < directions.Length && i < drifts.Length; i++)
            {
                if (!string.Equals(cases[i], combo, StringComparison.OrdinalIgnoreCase) ||
                    !directions[i].StartsWith(dir, StringComparison.OrdinalIgnoreCase))
                    continue;

                double value = Math.Abs(drifts[i]);
                if (!result.TryGetValue(stories[i], out var current) || value > current)
                    result[stories[i]] = value;
            }

            return result;
        }

        private static void AddMissingStories(
            cSapModel sap, List<EtabsHelper.StoryInfo> stories,
            Dictionary<string, double> displacements)
        {
            var known = new HashSet<string>(stories.Select(s => s.Name),
                StringComparer.OrdinalIgnoreCase);

            foreach (string storyName in displacements.Keys)
            {
                if (known.Contains(storyName)) continue;

                double elevation = 0.0, height = 0.0;
                if (sap.Story.GetElevation(storyName, ref elevation) != 0) continue;
                sap.Story.GetHeight(storyName, ref height);

                stories.Add(new EtabsHelper.StoryInfo
                {
                    Name = storyName,
                    Elevation = elevation,
                    Height = height
                });
                known.Add(storyName);
            }
        }
    }
}