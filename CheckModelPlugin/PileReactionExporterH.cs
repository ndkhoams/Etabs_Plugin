using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace Etabs_Ultimate_Tools
{
    /// <summary>
    /// Xuất kết quả kiểm tra phản lực cọc ra Excel: mỗi trường hợp tải = 1 sheet.
    /// Cột được dựng động: bỏ cột "SCT kéo" khi considerTension = false;
    /// bỏ các cột |FX|, |FY|, H, SCT ngang, KL ngang khi considerH = false.
    /// </summary>
    public static class PileReactionExporterH
    {
        private static readonly XLColor HeadFill = XLColor.FromArgb(197, 217, 241);

        public static void Export(string filePath, List<PileReactionCase> cases,
            bool considerTension, bool considerH)
        {
            using (var wb = new XLWorkbook())
            {
                if (cases != null)
                    foreach (var c in cases)
                        WriteCaseSheet(wb, c, considerTension, considerH);

                if (!wb.Worksheets.Any())
                    wb.Worksheets.Add("EMPTY");

                wb.SaveAs(filePath);
            }
        }

        // Danh sách khóa cột theo đúng thứ tự hiển thị (đồng bộ với bảng preview).
        private static List<string> BuildKeys(bool considerTension, bool considerH)
        {
            var keys = new List<string> { "Type", "Id", "Combo", "Reaction" };
            if (considerTension) keys.Add("TensCap");
            keys.Add("CompCap");
            keys.Add("Result");
            if (considerH)
            {
                keys.Add("Fx"); keys.Add("Fy"); keys.Add("H"); keys.Add("HCap"); keys.Add("HResult");
            }
            return keys;
        }

        private static string HeadOf(string key)
        {
            switch (key)
            {
                case "Type": return "Loại cọc";
                case "Id": return "Số hiệu cọc";
                case "Combo": return "Tổ hợp";
                case "Reaction": return "Phản lực đứng (kN)";
                case "TensCap": return "SCT kéo (kN)";
                case "CompCap": return "SCT nén (kN)";
                case "Result": return "KL đứng";
                case "Fx": return "|FX| (kN)";
                case "Fy": return "|FY| (kN)";
                case "H": return "H (kN)";
                case "HCap": return "SCT ngang (kN)";
                case "HResult": return "KL ngang";
            }
            return "";
        }

        private static double WidthOf(string key)
        {
            switch (key)
            {
                case "Type": return 11;
                case "Id": return 12;
                case "Combo": return 18;
                case "Reaction": return 16;
                case "TensCap": return 13;
                case "CompCap": return 13;
                case "Result": return 11;
                case "Fx": return 12;
                case "Fy": return 12;
                case "H": return 12;
                case "HCap": return 14;
                case "HResult": return 11;
            }
            return 12;
        }

        private static bool IsCenter(string key)
        {
            return key == "Type" || key == "Id" || key == "Combo"
                || key == "Result" || key == "HResult";
        }

        private static bool IsNum(string key)
        {
            return key == "Reaction" || key == "TensCap" || key == "CompCap"
                || key == "Fx" || key == "Fy" || key == "H" || key == "HCap";
        }

        private static void SetCell(IXLCell cell, string key, PileReactionRow row)
        {
            switch (key)
            {
                case "Type": cell.Value = row.PileType; break;
                case "Id": cell.Value = row.PileId; break;
                case "Combo": cell.Value = row.Combo; break;
                case "Reaction": cell.Value = Math.Round(row.Reaction, 1); break;
                case "TensCap": cell.Value = Math.Round(row.TensionCap, 1); break;
                case "CompCap": cell.Value = Math.Round(row.CompressionCap, 1); break;
                case "Result": cell.Value = row.Result; break;
                case "Fx": cell.Value = Math.Round(row.Fx, 1); break;
                case "Fy": cell.Value = Math.Round(row.Fy, 1); break;
                case "H": cell.Value = Math.Round(row.Horizontal, 1); break;
                case "HCap": cell.Value = Math.Round(row.HorizontalCap, 1); break;
                case "HResult": cell.Value = row.HResult; break;
            }
        }

        private static void WriteCaseSheet(XLWorkbook wb, PileReactionCase c,
            bool considerTension, bool considerH)
        {
            string sheetName = string.IsNullOrWhiteSpace(c.SheetName) ? "Sheet" : c.SheetName;
            var ws = wb.Worksheets.Add(sheetName);
            EtabsHelper.ApplyA4PageSetup(ws);

            var keys = BuildKeys(considerTension, considerH);
            int lastCol = keys.Count;

            ws.Cell("A1").Value = "KIỂM TRA KHẢ NĂNG CHỊ8U TẢI CỦA CỌC";
            ws.Range(1, 1, 1, lastCol).Merge();
            ws.Cell("A1").Style.Font.Bold = true;
            ws.Cell("A1").Style.Font.FontSize = 14;
            ws.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            ws.Cell("A2").Value = c.Title;
            ws.Range(2, 1, 2, lastCol).Merge();
            ws.Cell("A2").Style.Font.Bold = true;
            ws.Cell("A2").Style.Fill.BackgroundColor = HeadFill;
            ws.Cell("A2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            int headerRow = 3;
            for (int i = 0; i < keys.Count; i++)
                ws.Cell(headerRow, 1 + i).Value = HeadOf(keys[i]);
            StyleHeaderRange(ws.Range(headerRow, 1, headerRow, lastCol));

            int firstData = headerRow + 1;
            int r = firstData;
            var rows = c.Rows ?? new List<PileReactionRow>();
            foreach (var row in rows)
            {
                for (int i = 0; i < keys.Count; i++)
                {
                    string k = keys[i];
                    int col = 1 + i;
                    SetCell(ws.Cell(r, col), k, row);
                    if ((k == "Result" && IsFail(row.Result)) ||
                        (k == "HResult" && IsFail(row.HResult)))
                        ws.Cell(r, col).Style.Font.FontColor = XLColor.Red;
                }
                r++;
            }

            int lastData = Math.Max(firstData, r - 1);
            StyleBodyBox(ws.Range(firstData, 1, lastData, lastCol));

            for (int i = 0; i < keys.Count; i++)
            {
                string k = keys[i];
                int col = 1 + i;
                if (IsCenter(k))
                    ws.Range(firstData, col, lastData, col).Style.Alignment.Horizontal =
                        XLAlignmentHorizontalValues.Center;
                if (IsNum(k))
                    ws.Range(firstData, col, lastData, col).Style.NumberFormat.Format = "0";
                ws.Column(col).Width = WidthOf(k);
            }

            var used = ws.RangeUsed();
            if (used != null)
                used.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }

        private static bool IsFail(string result)
        {
            return !string.IsNullOrEmpty(result)
                && result.IndexOf("Không", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static void StyleHeaderRange(IXLRange header)
        {
            header.Style.Font.Bold = true;
            header.Style.Fill.BackgroundColor = HeadFill;
            header.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            header.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            header.Style.Alignment.WrapText = true;
            header.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            header.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        }

        private static void StyleBodyBox(IXLRange body)
        {
            body.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            body.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            body.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            body.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        }
    }
}
