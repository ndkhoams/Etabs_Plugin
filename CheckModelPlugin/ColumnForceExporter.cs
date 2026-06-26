using ETABSv1;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Etabs_Ultimate_Tools
{
    public enum ExportFileType
    {
        Text,
        Excel
    }

    public class ForceRow
    {
        public string Name { get; set; } = "";
        public double PU { get; set; }
        public double MUXT { get; set; }
        public double MUYT { get; set; }
        public double MUXB { get; set; }
        public double MUYB { get; set; }
    }

    public static class ColumnForceExporter
    {
        // Lay danh sach Load Combination tu model.
        public static List<string> GetCombos(cSapModel sap)
        {
            int numberNames = 0;
            string[] names = Array.Empty<string>();

            sap.RespCombo.GetNameList(ref numberNames, ref names);

            return names == null ? new List<string>() : names.ToList();
        }

        // Tinh noi luc cho cac cot/pier dang chon trong ETABS theo cac combo da chon.
        // columnCount/pierCount tra ve so luong da chon de form hien thi / kiem tra.
        public static List<ForceRow> Compute(cSapModel sap, IList<string> combos, out int columnCount, out int pierCount)
        {
            sap.SetPresentUnits(eUnits.kN_m_C);

            List<string> selectedColumns = GetSelectedColumns(sap);
            Dictionary<string, HashSet<string>> selectedPiers = GetSelectedPiers(sap);

            columnCount = selectedColumns.Count;
            pierCount = selectedPiers.Count;

            List<ForceRow> rows = new List<ForceRow>();

            if (combos == null || combos.Count == 0)
                return rows;

            if (columnCount == 0 && pierCount == 0)
                return rows;

            sap.Results.Setup.DeselectAllCasesAndCombosForOutput();

            foreach (string combo in combos)
                sap.Results.Setup.SetComboSelectedForOutput(combo);

            foreach (string column in selectedColumns)
                rows.AddRange(GetColumnForces(sap, column));

            foreach (var pierItem in selectedPiers)
                rows.AddRange(GetPierForces(sap, pierItem.Key, pierItem.Value));

            return rows;
        }

        public static void WriteText(List<ForceRow> rows, string path)
        {
            List<string> lines = new List<string>();

            foreach (ForceRow r in rows)
            {
                string line =
                    $"{r.Name}," +
                    $"{r.PU:0.##}," +
                    $"{r.MUXT:0.##}," +
                    $"{r.MUYT:0.##}," +
                    $"{r.MUXB:0.##}," +
                    $"{r.MUYB:0.##}";

                lines.Add(line);
            }

            string content = string.Join("\r\n", lines);

            System.IO.File.WriteAllText(path, content, Encoding.ASCII);
        }

        public static void WriteExcel(List<ForceRow> rows, string path)
        {
            using (XLWorkbook wb = new XLWorkbook())
            {
                IXLWorksheet ws = wb.Worksheets.Add("Results");

                ws.Cell(1, 1).Value = "NAME";
                ws.Cell(1, 2).Value = "PU";
                ws.Cell(1, 3).Value = "MUXT";
                ws.Cell(1, 4).Value = "MUYT";
                ws.Cell(1, 5).Value = "MUXB";
                ws.Cell(1, 6).Value = "MUYB";

                int row = 2;

                foreach (ForceRow r in rows)
                {
                    ws.Cell(row, 1).Value = r.Name;
                    ws.Cell(row, 2).Value = Math.Round(r.PU, 2);
                    ws.Cell(row, 3).Value = Math.Round(r.MUXT, 2);
                    ws.Cell(row, 4).Value = Math.Round(r.MUYT, 2);
                    ws.Cell(row, 5).Value = Math.Round(r.MUXB, 2);
                    ws.Cell(row, 6).Value = Math.Round(r.MUYB, 2);

                    row++;
                }

                ws.Columns().AdjustToContents();

                wb.SaveAs(path);
            }
        }

        private static List<string> GetSelectedColumns(cSapModel sap)
        {
            int numberItems = 0;
            int[] objectType = Array.Empty<int>();
            string[] objectName = Array.Empty<string>();

            sap.SelectObj.GetSelected(ref numberItems, ref objectType, ref objectName);

            List<string> columns = new List<string>();

            for (int i = 0; i < numberItems; i++)
            {
                if (objectType[i] != 2)
                    continue;

                string frameName = objectName[i];

                if (IsColumnByGeometry(sap, frameName))
                    columns.Add(frameName);
            }

            return columns;
        }

        private static Dictionary<string, HashSet<string>> GetSelectedPiers(cSapModel sap)
        {
            int numberItems = 0;
            int[] objectType = Array.Empty<int>();
            string[] objectName = Array.Empty<string>();

            sap.SelectObj.GetSelected(ref numberItems, ref objectType, ref objectName);

            Dictionary<string, HashSet<string>> piers =
                new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < numberItems; i++)
            {
                string pierName = "";
                string storyName = "";

                if (objectType[i] == 2)
                {
                    sap.FrameObj.GetPier(objectName[i], ref pierName);
                    storyName = GetFrameStoryFromLabelOrGeometry(sap, objectName[i]);
                }

                if (objectType[i] == 5)
                {
                    sap.AreaObj.GetPier(objectName[i], ref pierName);
                    storyName = GetAreaStoryFromLabel(sap, objectName[i]);
                }

                if (!string.IsNullOrWhiteSpace(pierName) &&
                    !pierName.Equals("None", StringComparison.OrdinalIgnoreCase) &&
                    !string.IsNullOrWhiteSpace(storyName))
                {
                    if (!piers.ContainsKey(pierName))
                        piers[pierName] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    piers[pierName].Add(storyName);
                }
            }

            return piers;
        }

        private static string GetFrameStoryFromLabelOrGeometry(cSapModel sap, string frameName)
        {
            string label = "";
            string storyName = "";

            int ret = sap.FrameObj.GetLabelFromName(frameName, ref label, ref storyName);

            if (ret == 0 && !string.IsNullOrWhiteSpace(storyName))
                return storyName.Trim();

            return GetColumnStory(sap, frameName);
        }

        private static string GetAreaStoryFromLabel(cSapModel sap, string areaName)
        {
            string label = "";
            string storyName = "";

            try
            {
                sap.AreaObj.GetLabelFromName(areaName, ref label, ref storyName);
            }
            catch
            {
                storyName = "";
            }

            return storyName;
        }

        private static bool IsColumnByGeometry(cSapModel sap, string frameName)
        {
            string p1 = "";
            string p2 = "";

            sap.FrameObj.GetPoints(frameName, ref p1, ref p2);

            double x1 = 0, y1 = 0, z1 = 0;
            double x2 = 0, y2 = 0, z2 = 0;

            sap.PointObj.GetCoordCartesian(p1, ref x1, ref y1, ref z1);
            sap.PointObj.GetCoordCartesian(p2, ref x2, ref y2, ref z2);

            double dz = Math.Abs(z2 - z1);
            double dh = Math.Sqrt(Math.Pow(x2 - x1, 2) + Math.Pow(y2 - y1, 2));

            return dz > dh;
        }

        private static List<ForceRow> GetColumnForces(cSapModel sap, string columnName)
        {
            List<ForceRow> rows = new List<ForceRow>();

            int numberResults = 0;

            string[] obj = Array.Empty<string>();
            double[] objSta = Array.Empty<double>();
            string[] elm = Array.Empty<string>();
            double[] elmSta = Array.Empty<double>();
            string[] loadCase = Array.Empty<string>();
            string[] stepType = Array.Empty<string>();
            double[] stepNum = Array.Empty<double>();

            double[] p = Array.Empty<double>();
            double[] v2 = Array.Empty<double>();
            double[] v3 = Array.Empty<double>();
            double[] t = Array.Empty<double>();
            double[] m2 = Array.Empty<double>();
            double[] m3 = Array.Empty<double>();

            sap.Results.FrameForce(
                columnName,
                eItemTypeElm.ObjectElm,
                ref numberResults,
                ref obj,
                ref objSta,
                ref elm,
                ref elmSta,
                ref loadCase,
                ref stepType,
                ref stepNum,
                ref p,
                ref v2,
                ref v3,
                ref t,
                ref m2,
                ref m3
            );

            if (numberResults == 0)
                return rows;

            string storyName = GetFrameStoryFromLabelOrGeometry(sap, columnName);

            var groups = Enumerable.Range(0, numberResults)
                .GroupBy(i => new
                {
                    Combo = loadCase[i],
                    StepType = stepType[i]
                });

            foreach (var g in groups)
            {
                int bottomIndex = g.OrderBy(i => objSta[i]).First();
                int topIndex = g.OrderByDescending(i => objSta[i]).First();

                string combo = loadCase[bottomIndex];
                string step = stepType[bottomIndex];

                string name;

                if (step.Contains("Single"))
                {
                    name = $"{storyName}-{columnName}-{combo}";
                }
                else if (step.Contains("Max"))
                {
                    name = $"{storyName}-{columnName}-{combo}Max";
                }
                else if (step.Contains("Min"))
                {
                    name = $"{storyName}-{columnName}-{combo}Min";
                }
                else
                {
                    name = $"{storyName}-{columnName}-{combo}-{step}";
                }

                rows.Add(new ForceRow
                {
                    Name = name,

                    // Quy doi dau giong file Excel mau CSI Column
                    PU = -p[bottomIndex],
                    MUXT = m2[topIndex],
                    MUYT = -m3[topIndex],
                    MUXB = -m2[bottomIndex],
                    MUYB = m3[bottomIndex]
                });
            }

            return rows;
        }

        private static List<ForceRow> GetPierForces(cSapModel sap, string pierName, HashSet<string> selectedStories)
        {
            List<ForceRow> rows = new List<ForceRow>();

            int numberResults = 0;

            string[] storyName = Array.Empty<string>();
            string[] pier = Array.Empty<string>();
            string[] loadCase = Array.Empty<string>();
            string[] location = Array.Empty<string>();

            double[] p = Array.Empty<double>();
            double[] v2 = Array.Empty<double>();
            double[] v3 = Array.Empty<double>();
            double[] t = Array.Empty<double>();
            double[] m2 = Array.Empty<double>();
            double[] m3 = Array.Empty<double>();

            sap.Results.PierForce(
                ref numberResults,
                ref storyName,
                ref pier,
                ref loadCase,
                ref location,
                ref p,
                ref v2,
                ref v3,
                ref t,
                ref m2,
                ref m3
            );

            if (numberResults == 0)
                return rows;

            var groups = Enumerable.Range(0, numberResults)
                .Where(i =>
                    pier[i].Equals(pierName, StringComparison.OrdinalIgnoreCase) &&
                    selectedStories.Contains(storyName[i]))
                .GroupBy(i => new
                {
                    Story = storyName[i],
                    Pier = pier[i],
                    Combo = loadCase[i]
                });

            foreach (var g in groups)
            {
                List<int> bottomIndices = g
                    .Where(i => !IsTopLocation(location[i]))
                    .ToList();

                List<int> topIndices = g
                    .Where(i => IsTopLocation(location[i]))
                    .ToList();

                if (bottomIndices.Count == 0)
                    bottomIndices = g.ToList();

                if (topIndices.Count == 0)
                    topIndices = g.ToList();

                if (bottomIndices.Count <= 1 && topIndices.Count <= 1)
                {
                    int bottomIndex = bottomIndices
                        .OrderBy(i => IsTopLocation(location[i]) ? 1 : 0)
                        .First();

                    int topIndex = topIndices
                        .OrderByDescending(i => IsTopLocation(location[i]) ? 1 : 0)
                        .First();

                    string name =
                        $"{storyName[bottomIndex]}-{pierName}-{loadCase[bottomIndex]}";

                    rows.Add(CreatePierForceRow(name, bottomIndex, topIndex, p, m2, m3));
                }
                else
                {
                    int bottomMaxIndex = bottomIndices.OrderByDescending(i => p[i]).First();
                    int bottomMinIndex = bottomIndices.OrderBy(i => p[i]).First();

                    int topMaxIndex = topIndices.OrderByDescending(i => p[i]).First();
                    int topMinIndex = topIndices.OrderBy(i => p[i]).First();

                    string baseName =
                        $"{storyName[bottomMaxIndex]}-{pierName}-{loadCase[bottomMaxIndex]}";

                    rows.Add(CreatePierForceRow(baseName + "Max", bottomMaxIndex, topMaxIndex, p, m2, m3));

                    if (bottomMinIndex != bottomMaxIndex || topMinIndex != topMaxIndex)
                        rows.Add(CreatePierForceRow(baseName + "Min", bottomMinIndex, topMinIndex, p, m2, m3));
                }
            }

            return rows;
        }

        private static ForceRow CreatePierForceRow(
            string name,
            int bottomIndex,
            int topIndex,
            double[] p,
            double[] m2,
            double[] m3)
        {
            return new ForceRow
            {
                Name = name,

                // Quy doi dau giong file Excel mau CSI Column
                PU = -p[bottomIndex],
                MUXT = m2[topIndex],
                MUYT = -m3[topIndex],
                MUXB = -m2[bottomIndex],
                MUYB = m3[bottomIndex]
            };
        }

        private static string GetColumnStory(cSapModel sap, string columnName)
        {
            string p1 = "";
            string p2 = "";

            sap.FrameObj.GetPoints(columnName, ref p1, ref p2);

            double x1 = 0, y1 = 0, z1 = 0;
            double x2 = 0, y2 = 0, z2 = 0;

            sap.PointObj.GetCoordCartesian(p1, ref x1, ref y1, ref z1);
            sap.PointObj.GetCoordCartesian(p2, ref x2, ref y2, ref z2);

            double zMid = (z1 + z2) / 2.0;

            int numberStories = 0;
            string[] storyNames = Array.Empty<string>();
            double[] storyElevations = Array.Empty<double>();
            double[] storyHeights = Array.Empty<double>();
            bool[] isMasterStory = Array.Empty<bool>();
            string[] similarToStory = Array.Empty<string>();
            bool[] spliceAbove = Array.Empty<bool>();
            double[] spliceHeight = Array.Empty<double>();

            sap.Story.GetStories(
                ref numberStories,
                ref storyNames,
                ref storyElevations,
                ref storyHeights,
                ref isMasterStory,
                ref similarToStory,
                ref spliceAbove,
                ref spliceHeight
            );

            if (numberStories == 0)
                return "UnknownStory";

            int nearest = 0;
            double minDiff = Math.Abs(storyElevations[0] - zMid);

            for (int i = 1; i < numberStories; i++)
            {
                double diff = Math.Abs(storyElevations[i] - zMid);

                if (diff < minDiff)
                {
                    minDiff = diff;
                    nearest = i;
                }
            }

            return storyNames[nearest];
        }

        private static bool IsTopLocation(string location)
        {
            if (string.IsNullOrWhiteSpace(location))
                return false;

            return location.IndexOf("Top", StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
