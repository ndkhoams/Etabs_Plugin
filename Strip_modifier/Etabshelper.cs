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

            sap.Story.GetStories(ref n, ref names, ref elevations, ref heights,
                ref isMaster, ref similarTo, ref spliceAbove, ref spliceHeight);

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

        /// <summary>Trả về true nếu tên tầng là "Base" hoặc chứa "Base".</summary>
        public static bool IsBaseLevel(string storyName)
        {
            if (string.IsNullOrWhiteSpace(storyName)) return false;
            return storyName.Trim().IndexOf("Base", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        /// <summary>outputCase trùng selectedName (rỗng, bằng, hoặc chứa nhau).</summary>
        public static bool IsSameOrBlank(string outputCase, string selectedName)
        {
            if (string.IsNullOrWhiteSpace(outputCase)) return true;
            if (string.Equals(outputCase.Trim(), selectedName.Trim(), StringComparison.OrdinalIgnoreCase))
                return true;
            return outputCase.IndexOf(selectedName, StringComparison.OrdinalIgnoreCase) >= 0
                || selectedName.IndexOf(outputCase, StringComparison.OrdinalIgnoreCase) >= 0;
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
