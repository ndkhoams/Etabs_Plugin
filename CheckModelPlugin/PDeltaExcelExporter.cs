using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace CheckModelPlugin
{
    public static class PDeltaExcelExporter
    {
        public static void Export(string filePath, List<PDeltaCheckRow> rows, double qFactor,
            List<TopDisplacementRow> windRows = null)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("P-DELTA");
                EtabsHelper.ApplyA4PageSetup(ws);

                WriteHeaderAndNotes(ws);

                int row = 18;
                WriteDirectionSection(ws,
                    rows.Where(x => x.Direction.Equals("X", StringComparison.OrdinalIgnoreCase)).ToList(),
                    "2. Kiểm tra theo phương X", qFactor, ref row);

                row += 2;
                WriteDirectionSection(ws,
                    rows.Where(x => x.Direction.Equals("Y", StringComparison.OrdinalIgnoreCase)).ToList(),
                    "3. Kiểm tra theo phương Y", qFactor, ref row);

                ws.Column(1).Width = 3;
                for (int c = 2; c <= 9; c++) ws.Column(c).Width = 12;
                ws.Rows().Height = 18;
                ws.Row(1).Height = 21;
                ws.Row(2).Height = 18;

                ws.SheetView.View = XLSheetViewOptions.Normal;
                ws.SheetView.FreezeRows(0);
                ws.RangeUsed().Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                if (windRows != null && windRows.Count > 0)
                    WriteWindSheet(wb, windRows);

                wb.SaveAs(filePath);
            }
        }

        // ─── Style helpers (gom style header/body lặp lại) ───────────────
        private static void StyleHeaderRange(IXLRange header)
        {
            header.Style.Font.Bold = true;
            header.Style.Fill.BackgroundColor = XLColor.LightGray;
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

        private static void WriteHeaderAndNotes(IXLWorksheet ws)
        {
            ws.Cell("A1").Value = "KIỂM TRA ĐIỀU KIỆN P-DELTA";
            ws.Range("A1:I1").Merge();
            ws.Cell("A1").Style.Font.Bold = true;
            ws.Cell("A1").Style.Font.FontSize = 14;
            ws.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            ws.Cell("A2").Value = "(Theo TCVN 9386-1:2025)";
            ws.Range("A2:I2").Merge();
            ws.Cell("A2").Style.Font.Italic = true;
            ws.Cell("A2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            ws.Cell("A3").Value = "1. Chú thích";
            ws.Cell("A3").Style.Font.Bold = true;

            ws.Cell("B4").Value = "θ - hệ số độ nhạy của chuyển vị tương đối giữa các tầng";
            ws.Cell("B5").Value = "θ = q × drift × Ptot / Vtot";
            ws.Range("B5:I5").Merge();
            ws.Cell("B5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell("B5").Style.Font.FontSize = 10;

            ws.Cell("B6").Value = "Ptot - Tổng tải trọng đứng (trọng lực) tại tầng đang xét và tất cả các tầng bên trên nó";
            ws.Range("B6:I6").Merge();
            ws.Cell("B7").Value = "trong tình huống thiết kế động đất";
            ws.Cell("B8").Value = "Vtot - tổng lực cắt tầng do động đất gây ra";
            ws.Cell("B9").Value = "drift - tỷ số lệch tầng đàn hồi, lấy từ ETABS Story Drift";
            ws.Cell("B10").Value = "q × drift - tỷ số lệch tầng thiết kế";
            ws.Cell("B11").Value = "q - hệ số ứng xử";
            ws.Range("B11:I11").Merge();

            ws.Cell("A13").Value = "Các điều kiện khống chế:";
            ws.Cell("A13").Style.Font.Bold = true;
            ws.Cell("B14").Value = "θ ≤ 0.10 : không cần xét tới hiệu ứng bậc 2";
            ws.Cell("B15").Value = "0.10 < θ ≤ 0.20 : xét gần đúng bằng cách nhân các hệ quả tác động với 1/(1-θ)";
            ws.Range("B15:I15").Merge();
            ws.Cell("B16").Value = "θ không được vượt quá 0.30";
        }

        private static void WriteDirectionSection(IXLWorksheet ws, List<PDeltaCheckRow> dirRows,
            string sectionTitle, double qFactor, ref int r)
        {
            ws.Cell(r, 1).Value = sectionTitle;
            ws.Cell(r, 1).Style.Font.Bold = true;
            ws.Range(r, 1, r, 9).Merge();
            r++;

            double thetaMax = dirRows.Count > 0 ? dirRows.Max(x => x.Theta) : 0.0;

            ws.Cell(r, 2).Value = "Hệ số ứng xử :";
            ws.Cell(r, 4).Value = "q =";
            ws.Cell(r, 5).Value = qFactor;
            ws.Cell(r, 5).Style.NumberFormat.Format = "0.###";
            r++;

            ws.Cell(r, 2).Value = "Hệ số độ nhạy :";
            ws.Cell(r, 4).Value = "θmax =";
            ws.Cell(r, 5).Value = thetaMax;
            ws.Cell(r, 5).Style.NumberFormat.Format = "0.0000";
            r++;

            ws.Cell(r, 2).Value = "Kết luận:";
            ws.Cell(r, 3).Value = GetSummary(thetaMax);
            ws.Range(r, 3, r, 9).Merge();
            ws.Cell(r, 3).Style.Font.Bold = true;
            ws.Cell(r, 3).Style.Font.Italic = true;
            r++;

            int headerRow1 = r;
            int headerRow2 = r + 1;

            string[] heads = { "Tầng", "drift", "q × drift", "Ptot", "Vtot", "θ", "1/(1-θ)", "Kiểm tra" };
            for (int i = 0; i < heads.Length; i++)
                ws.Cell(headerRow1, 2 + i).Value = heads[i];

            ws.Cell(headerRow2, 5).Value = "(kN)";
            ws.Cell(headerRow2, 6).Value = "(kN)";

            foreach (int col in new[] { 2, 3, 4, 7, 8, 9 })
                ws.Range(headerRow1, col, headerRow2, col).Merge();

            StyleHeaderRange(ws.Range(headerRow1, 2, headerRow2, 9));

            r += 2;
            int firstData = r;

            foreach (var x in dirRows.OrderByDescending(x => x.Elevation))
            {
                ws.Cell(r, 2).Value = x.Story;
                ws.Cell(r, 3).Value = x.ElasticDrift;
                ws.Cell(r, 4).Value = x.DesignDrift;
                ws.Cell(r, 5).Value = x.Ptot;
                ws.Cell(r, 6).Value = x.Vtot;
                ws.Cell(r, 7).Value = x.Theta;
                ws.Cell(r, 8).Value = x.Amplification;
                ws.Cell(r, 9).Value = ShortCheck(x.Theta);
                r++;
            }

            int lastData = Math.Max(firstData, r - 1);
            StyleBodyBox(ws.Range(firstData, 2, lastData, 9));
            ws.Range(firstData, 2, lastData, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            ws.Range(firstData, 7, lastData, 8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Range(firstData, 9, lastData, 9).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

            ws.Range(firstData, 3, lastData, 4).Style.NumberFormat.Format = "0.00000";   // drift, q × drift
            ws.Range(firstData, 5, lastData, 6).Style.NumberFormat.Format = "0";         // Ptot, Vtot
            ws.Range(firstData, 7, lastData, 8).Style.NumberFormat.Format = "0.000";     // θ, 1/(1-θ)
        }

        private static void WriteWindSheet(XLWorkbook wb, List<TopDisplacementRow> rows)
        {
            var ws = wb.Worksheets.Add("CHUYEN VI DINH");
            EtabsHelper.ApplyA4PageSetup(ws);

            var validRows = (rows ?? new List<TopDisplacementRow>())
                .Where(x => x.TopElevation > 1e-9 && !EtabsHelper.IsBaseLevel(x.TopStory))
                .ToList();

            var xRows = BuildMaxByStory(validRows, "X");
            var yRows = BuildMaxByStory(validRows, "Y");

            var storyRows = validRows
                .GroupBy(x => x.TopStory, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.OrderByDescending(r => r.TopElevation).First())
                .OrderByDescending(x => x.StoryElevation)
                .ToList();

            double limit = validRows.Count > 0 ? validRows[0].LimitDenominator : 500.0;
            if (limit <= 0) limit = 500.0;
            string limitText = limit.ToString("0");

            string comboX = xRows.Count > 0 ? xRows.Values.First().Combo : "";
            string comboY = yRows.Count > 0 ? yRows.Values.First().Combo : "";
            string comboText = string.Equals(comboX, comboY, StringComparison.OrdinalIgnoreCase)
                ? comboX
                : "X: " + comboX + "; Y: " + comboY;

            // Tính 1 lần duy nhất, dùng lại cho cả kết luận lẫn bảng
            var computed = storyRows.Select(st =>
            {
                xRows.TryGetValue(st.TopStory, out var xr);
                yRows.TryGetValue(st.TopStory, out var yr);
                double dx = xr != null ? xr.TopDisplacementMm : 0.0;
                double dy = yr != null ? yr.TopDisplacementMm : 0.0;
                double limitMm = st.TopElevation * 1000.0 / limit;
                return new { Story = st, Dx = dx, Dy = dy, LimitMm = limitMm, Ok = dx <= limitMm && dy <= limitMm };
            }).ToList();

            bool anyNg = computed.Any(c => !c.Ok);

            ws.Cell("A2").Value = "KIỂM TRA CHUYỂN VỊ ĐỈNH CÔNG TRÌNH";
            ws.Range("A2:H2").Merge();
            ws.Cell("A2").Style.Font.Bold = true;
            ws.Cell("A2").Style.Font.FontSize = 14;
            ws.Cell("A2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            ws.Cell("A3").Value = "(Theo TCVN 2737:2023)";
            ws.Range("A3:H3").Merge();
            ws.Cell("A3").Style.Font.Italic = true;
            ws.Cell("A3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            ws.Cell("A5").Value = "1. Cơ sở lý thuyết";
            ws.Cell("A5").Style.Font.Bold = true;

            ws.Cell("B6").Value = "Theo TCVN 2737:2023, mọi cấu kiện xây dựng phải thỏa mãn điều kiện: f ≤ fu";
            ws.Range("B6:H6").Merge();
            ws.Cell("B7").Value = "trong đó f là chuyển vị tính toán, còn fu là giá trị giới hạn quy định trong tiêu chuẩn.";
            ws.Range("B7:H7").Merge();
            ws.Cell("B8").Value = "Đối với nhà nhiều tầng: Chuyển vị ngang tổng thể giới hạn là H/" + limitText;
            ws.Range("B8:H8").Merge();
            ws.Cell("B9").Value = "      max [Δ] < H/" + limitText;
            ws.Range("B9:H9").Merge();
            ws.Cell("B10").Value = "Trong đó:";
            ws.Cell("C10").Value = "Δ là chuyển vị ngang";
            ws.Range("C10:H10").Merge();
            ws.Cell("C11").Value = "H được tính là khoảng cách từ mặt móng đến trục của xà đỡ mái.";
            ws.Range("C11:H11").Merge();

            ws.Cell("A12").Value = "2. Kiểm tra chuyển vị";
            ws.Cell("A12").Style.Font.Bold = true;

            ws.Cell("B13").Value = "Tổ hợp kiểm tra:";
            ws.Cell("D13").Value = comboText;
            ws.Range("D13:H13").Merge();

            ws.Cell("B14").Value = "Kết luận:";
            ws.Cell("C14").Value = anyNg
                ? "Công trình không đảm bảo điều kiện chuyển vị đỉnh."
                : "Công trình đảm bảo điều kiện chuyển vị đỉnh.";
            ws.Range("C14:H14").Merge();
            ws.Cell("C14").Style.Font.Bold = true;
            ws.Cell("C14").Style.Font.Italic = true;

            int headerRow = 15;
            string[] heads = { "Tầng", "Cao độ tầng\n(m)", "H\n(m)", "ΔX\n(mm)", "ΔY\n(mm)", "H/" + limitText + "\n(mm)", "Kiểm tra" };
            for (int i = 0; i < heads.Length; i++)
                ws.Cell(headerRow, 2 + i).Value = heads[i];

            StyleHeaderRange(ws.Range(headerRow, 2, headerRow, 8));

            int r = headerRow + 1;
            int firstData = r;
            foreach (var c in computed)
            {
                ws.Cell(r, 2).Value = c.Story.TopStory;
                ws.Cell(r, 3).Value = c.Story.StoryElevation;
                ws.Cell(r, 4).Value = c.Story.TopElevation;
                ws.Cell(r, 5).Value = c.Dx;
                ws.Cell(r, 6).Value = c.Dy;
                ws.Cell(r, 7).Value = c.LimitMm;
                ws.Cell(r, 8).Value = c.Ok ? "OK" : "NG";
                r++;
            }

            int lastData = Math.Max(firstData, r - 1);
            if (lastData >= firstData)
            {
                StyleBodyBox(ws.Range(firstData, 2, lastData, 8));
                ws.Range(firstData, 2, lastData, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                ws.Range(firstData, 8, lastData, 8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Range(firstData, 3, lastData, 3).Style.NumberFormat.Format = "+0.000;-0.000;0.000";
                ws.Range(firstData, 4, lastData, 4).Style.NumberFormat.Format = "0.000";
                ws.Range(firstData, 5, lastData, 6).Style.NumberFormat.Format = "0.0";
                ws.Range(firstData, 7, lastData, 7).Style.NumberFormat.Format = "0";
            }

            ws.Column(1).Width = 3;
            ws.Column(2).Width = 14;
            ws.Column(3).Width = 16;
            ws.Column(4).Width = 13;
            for (int col = 5; col <= 8; col++) ws.Column(col).Width = 14;
            ws.Rows().Height = 18;
            ws.Row(1).Height = 21;
            ws.Row(2).Height = 21;
            ws.Row(3).Height = 18;
            ws.Row(15).Height = 30;

            ws.SheetView.View = XLSheetViewOptions.Normal;
            ws.RangeUsed().Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }

        // Gom logic GroupBy + lấy dòng có |chuyển vị| lớn nhất (trước đây lặp cho X và Y)
        private static Dictionary<string, TopDisplacementRow> BuildMaxByStory(
            List<TopDisplacementRow> rows, string direction)
        {
            return rows
                .Where(x => x.Direction.Equals(direction, StringComparison.OrdinalIgnoreCase))
                .GroupBy(x => x.TopStory, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(
                    g => g.Key,
                    g => g.OrderByDescending(r => Math.Abs(r.TopDisplacement)).First(),
                    StringComparer.OrdinalIgnoreCase);
        }

        private static string ShortCheck(double theta)
        {
            if (theta <= 0.10) return "OK";
            if (theta <= 0.20) return "> 0.1";
            if (theta <= 0.30) return "> 0.2";
            return "NG";
        }

        private static string GetSummary(double theta)
        {
            if (theta <= 0.10) return "Không cần xét đến các hiệu ứng bậc 2";
            if (theta <= 0.20) return "Cần xét gần đúng hiệu ứng bậc 2 bằng hệ số 1/(1-θ)";
            if (theta <= 0.30) return "Ảnh hưởng P-Delta lớn, cần kiểm tra chính xác hơn";
            return "Không đạt điều kiện θ ≤ 0.30, cần điều chỉnh thiết kế";
        }
    }
}