using ETABSv1;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Etabs_Ultimate_Tools
{
    /// <summary>
    /// Các hàm tiện ích dùng chung, tránh trùng lặp code giữa các class.
    /// </summary>
    internal static class EtabsHelper
    {
        public class StoryInfo
        {
            public string Name { get; set; }
            public double Elevation { get; set; }
            public double Height { get; set; }
        }

        /// <summary>Đọc danh sách tầng từ ETABS, sắp xếp theo cao độ tăng dần.</summary>
        public static List<StoryInfo> ReadStories(cSapModel sap)
        {
            int n = 0;
            string[] names = null;
            double[] elevations = null, heights = null;
            bool[] isMaster = null, spliceAbove = null;
            string[] similarTo = null;
            double[] spliceHeight = null;

            int ret = sap.Story.GetStories(ref n, ref names, ref elevations, ref heights,
                ref isMaster, ref similarTo, ref spliceAbove, ref spliceHeight);
            if (ret != 0 || n < 0 || names == null || elevations == null ||
                names.Length < n || elevations.Length < n)
                throw new InvalidOperationException(
                    "ETABS Story.GetStories thất bại hoặc trả về mảng tầng không hợp lệ (return code " + ret + ").");

            var list = new List<StoryInfo>(n);
            for (int i = 0; i < n; i++)
                list.Add(new StoryInfo
                {
                    Name = names[i],
                    Elevation = elevations[i],
                    Height = heights != null && i < heights.Length ? heights[i] : 0.0
                });

            return list.OrderBy(x => x.Elevation).ToList();
        }

        /// <summary>Bổ sung tầng có trong kết quả Story Drifts nhưng thiếu trong GetStories.</summary>
        public static void AddMissingStoriesFromDrifts(
            cSapModel sap, List<StoryInfo> stories, params string[] loadCases)
        {
            var knownStories = new HashSet<string>(
                stories.Select(s => s.Name), StringComparer.OrdinalIgnoreCase);

            foreach (string loadCase in loadCases
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                try
                {
                    SelectCaseOrCombo(sap, loadCase);

                    int count = 0;
                    string[] storyNames = null, caseNames = null, stepTypes = null;
                    string[] directions = null, labels = null;
                    double[] stepNumbers = null, drifts = null, x = null, y = null, z = null;
                    int ret = sap.Results.StoryDrifts(ref count, ref storyNames, ref caseNames,
                        ref stepTypes, ref stepNumbers, ref directions, ref drifts,
                        ref labels, ref x, ref y, ref z);
                    if (ret != 0 || storyNames == null) continue;

                    for (int i = 0; i < count && i < storyNames.Length; i++)
                    {
                        string storyName = storyNames[i];
                        if (string.IsNullOrWhiteSpace(storyName) || knownStories.Contains(storyName))
                            continue;
                        if (caseNames != null && i < caseNames.Length &&
                            !string.Equals(caseNames[i], loadCase, StringComparison.OrdinalIgnoreCase))
                            continue;

                        double elevation = 0.0, height = 0.0;
                        if (sap.Story.GetElevation(storyName, ref elevation) != 0) continue;
                        sap.Story.GetHeight(storyName, ref height);

                        stories.Add(new StoryInfo
                        {
                            Name = storyName,
                            Elevation = elevation,
                            Height = height
                        });
                        knownStories.Add(storyName);
                    }
                }
                catch
                {
                    // Keep the original story list if this result set is unavailable.
                }
            }

            stories.Sort((a, b) => a.Elevation.CompareTo(b.Elevation));
        }

        /// <summary>Trả về true nếu tên tầng là "Base" hoặc chứa "Base".</summary>
        public static bool IsBaseLevel(string storyName)
        {
            if (string.IsNullOrWhiteSpace(storyName)) return false;
            return storyName.Trim().IndexOf("Base", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        /// <summary>So sánh tên output case sau khi trim, không phân biệt hoa thường.</summary>
        public static bool IsSameOrBlank(string outputCase, string selectedName)
        {
            return !string.IsNullOrWhiteSpace(outputCase)
                && !string.IsNullOrWhiteSpace(selectedName)
                && string.Equals(outputCase.Trim(), selectedName.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        public static bool IsSameOrEnvelopeCase(string outputCase, string selectedName)
        {
            if (IsSameOrBlank(outputCase, selectedName)) return true;
            if (string.IsNullOrWhiteSpace(outputCase) || string.IsNullOrWhiteSpace(selectedName))
                return false;

            string actual = outputCase.Trim();
            string selected = selectedName.Trim();
            string[] suffixes = { " (Max)", " (Min)", "_max", "_min", "-max", "-min", " Max", " Min" };
            foreach (string suffix in suffixes)
            {
                if (actual.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(actual.Substring(0, actual.Length - suffix.Length).TrimEnd(),
                        selected, StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            return false;
        }

        public static void SelectCaseOrCombo(cSapModel sap, string name)
        {
            sap.Results.Setup.DeselectAllCasesAndCombosForOutput();
            int ret = sap.Results.Setup.SetCaseSelectedForOutput(name);
            if (ret != 0) sap.Results.Setup.SetComboSelectedForOutput(name);
        }

        public static void SelectComboOnly(cSapModel sap, string comboName)
        {
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

        /// <summary>Áp dụng thiết lập trang in chuẩn A4 cho một worksheet.</summary>
        public static void ApplyA4PageSetup(ClosedXML.Excel.IXLWorksheet ws)
        {
            ws.PageSetup.PaperSize = ClosedXML.Excel.XLPaperSize.A4Paper;
            ws.PageSetup.PageOrientation = ClosedXML.Excel.XLPageOrientation.Portrait;
            ws.PageSetup.FitToPages(1, 0);
            ws.PageSetup.Margins.Left = 0.75;
            ws.PageSetup.Margins.Right = 0.75;
            ws.PageSetup.Margins.Top = 0.75;
            ws.PageSetup.Margins.Bottom = 0.50;
            ws.PageSetup.Margins.Header = 0.50;
            ws.PageSetup.Margins.Footer = 0.75;
            ws.PageSetup.CenterHorizontally = true;

            ws.Style.Font.FontName = "Arial";
            ws.Style.Font.FontSize = 11;
        }
    }
}
