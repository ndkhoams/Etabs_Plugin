using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace Etabs_Ultimate_Tools
{
    /// <summary>
    /// Xuất kết quả kiểm tra phản lực cọc CÓ LỰC NGANG ra Excel:
    /// mỗi trường hợp tải = 1 sheet, có thêm cột |FX|, |FY|, H và SCT ngang.
    /// </summary>
    public static class PileReactionExporterH
    {
        private static readonly XLColor HeadFill = XLColor.FromArgb(197, 217, 241);

        public static void Export(string filePath, List<PileReactionCase> cases)
        {
            using (var wb = new XLWorkbook())
            {
                if (cases != null)
                    foreach (var c in cases)
                        WriteCaseSheet(wb, c);

                if (!wb.Worksheets.Any())
                    wb.Worksheets.Add("EMPTY");

                wb.SaveAs(filePath);
            }
        }

        private static void WriteCaseSheet(XLWorkbook wb, PileReactionCase c)
        {
            string sheetName = string.IsNullOrWhiteSpace(c.SheetName) ? "Sheet" : c.SheetName;
            var ws = wb.Worksheets.Add(sheetName);
            EtabsHelper.ApplyA4PageSetup(ws);

            ws.Cell("A1").Value = "KIỂM TRA KHẢ NĂNG CHỊU TẢI CỦA CỌC (CÓ LỰC NGANG)";
            ws.Range("A1:L1").Merge();
            ws.Cell("A1").Style.Font.Bold = true;
            ws.Cell("A1").Style.Font.FontSize = 14;
            ws.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            ws.Cell("A2").Value = c.Title;
            ws.Range("A2:L2").Merge();
            ws.Cell("A2").Style.Font.Bold = true;
            ws.Cell("A2").Style.Fill.BackgroundColor = HeadFill;
            ws.Cell("A2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            int headerRow = 3;
            string[] heads = { "Loại cọc", "Số hiệu cọc", "Tổ hợp", "Phản lực đứng (kN)",
                "SCT kéo (kN)", "SCT nén (kN)", "KL đứng",
                "|FX| (kN)", "|FY| (kN)", "H (kN)", "SCT ngang (kN)", "KL ngang" };
            for (int i = 0; i < heads.Length; i++)
                ws.Cell(headerRow, 1 + i).Value = heads[i];
            StyleHeaderRange(ws.Range(headerRow, 1, headerRow, 12));

            int firstData = headerRow + 1;
            int r = firstData;
            var rows = c.Rows ?? new List<PileReactionRow>();
            foreach (var row in rows)
            {
                ws.Cell(r, 1).Value = row.PileType;
                ws.Cell(r, 2).Value = row.PileId;
                ws.Cell(r, 3).Value = row.Combo;
                ws.Cell(r, 4).Value = Math.Round(row.Reaction, 1);
                ws.Cell(r, 5).Value = Math.Round(row.TensionCap, 1);
                ws.Cell(r, 6).Value = Math.Round(row.CompressionCap, 1);
                ws.Cell(r, 7).Value = row.Result;
                ws.Cell(r, 8).Value = Math.Round(row.Fx, 1);
                ws.Cell(r, 9).Value = Math.Round(row.Fy, 1);
                ws.Cell(r, 10).Value = Math.Round(row.Horizontal, 1);
                ws.Cell(r, 11).Value = Math.Round(row.HorizontalCap, 1);
                ws.Cell(r, 12).Value = row.HResult;

                if (IsFail(row.Result)) ws.Cell(r, 7).Style.Font.FontColor = XLColor.Red;
                if (IsFail(row.HResult)) ws.Cell(r, 12).Style.Font.FontColor = XLColor.Red;
                r++;
            }

            int lastData = Math.Max(firstData, r - 1);
            StyleBodyBox(ws.Range(firstData, 1, lastData, 12));
            ws.Range(firstData, 1, lastData, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Range(firstData, 7, lastData, 7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Range(firstData, 12, lastData, 12).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Range(firstData, 4, lastData, 6).Style.NumberFormat.Format = "0";
            ws.Range(firstData, 8, lastData, 11).Style.NumberFormat.Format = "0";

            ws.Column(1).Width = 11;
            ws.Column(2).Width = 12;
            ws.Column(3).Width = 18;
            ws.Column(4).Width = 16;
            ws.Column(5).Width = 13;
            ws.Column(6).Width = 13;
            ws.Column(7).Width = 11;
            ws.Column(8).Width = 12;
            ws.Column(9).Width = 12;
            ws.Column(10).Width = 12;
            ws.Column(11).Width = 14;
            ws.Column(12).Width = 11;

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
