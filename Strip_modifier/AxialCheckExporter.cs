using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace Etabs_Ultimate_Tools
{
    public static class AxialCheckExporter
    {
        // ── Hằng số cột ──────────────────────────────────────────────────────────
        private const int ColSTT = 1;
        private const int ColTang = 2;
        private const int ColLoai = 3;
        private const int ColLabel = 4;
        private const int ColCombo = 5;
        private const int ColNed = 6;
        private const int ColT3 = 7;
        private const int ColT2 = 8;
        private const int ColAc = 9;
        private const int ColAcFcd = 10;
        private const int ColVd = 11;
        private const int ColVdLimit = 12;
        private const int ColKetLuan = 13;
        private const int TotalCols = 13;

        private static readonly string[] ConcreteGrades =
            { "B15", "B20", "B22.5", "B25", "B30", "B35", "B40",
              "B45", "B50", "B55", "B60", "B70", "B80" };

        // ── Entry point ──────────────────────────────────────────────────────────
        public static void Export(
            IReadOnlyList<AxialCheckRow> rows,
            string filePath,
            double alphaCc,
            double gammaC,
            double colLimit,
            double wallLimit)
        {
            if (rows == null || rows.Count == 0)
                throw new InvalidOperationException("Không có dữ liệu để xuất Excel.");

            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Axial Check");

                SetupPage(ws);
                WriteTitle(ws);
                WriteConcreteBlock(ws, rows[0].FckCube);
                WriteCoefficients(ws, alphaCc, gammaC);
                WriteConditions(ws, colLimit, wallLimit);
                WriteOverallResult(ws);
                WriteTableHeader(ws);
                WriteData(ws, rows, colLimit, wallLimit);
                ApplyTableStyle(ws, rows.Count);

                wb.CalculateMode = XLCalculateMode.Auto;
                wb.SaveAs(filePath);
            }
        }

        // ── Page setup ───────────────────────────────────────────────────────────
        private static void SetupPage(IXLWorksheet ws)
        {
            ws.PageSetup.PaperSize = XLPaperSize.A4Paper;
            ws.PageSetup.PageOrientation = XLPageOrientation.Portrait;
            ws.PageSetup.FitToPages(1, 0);
            ws.PageSetup.SetRowsToRepeatAtTop(1, HeaderRow);
            ws.SheetView.View = XLSheetViewOptions.PageBreakPreview;
            ws.SheetView.ZoomScale = 130;
        }

        // ── Tiêu đề ──────────────────────────────────────────────────────────────
        private static void WriteTitle(IXLWorksheet ws)
        {
            var range = ws.Range("A2:M2");
            range.Merge();
            range.FirstCell().Value = "KIỂM TRA HỆ SỐ LỰC DỌC QUY ĐỔI";
            range.Style.Font.Bold = true;
            range.Style.Font.FontSize = 18;
            range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }

        // ── Khối bê tông ─────────────────────────────────────────────────────────
        private static void WriteConcreteBlock(IXLWorksheet ws, double fckCube)
        {
            SetBoldUnderline(ws.Cell("A4"), "Bê tông:");

            string grade = $"B{fckCube:0.#}";
            var gradeCell = ws.Cell("C4");
            gradeCell.Value = grade;
            gradeCell.Style.Font.Bold = true;
            gradeCell.Style.Font.FontColor = XLColor.Blue;
            gradeCell.Style.Fill.BackgroundColor = XLColor.FromHtml("#EAF2F8");
            ws.Cell("C4").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

            // Dropdown cấp bền bê tông
            var dv = gradeCell.CreateDataValidation();
            dv.IgnoreBlanks = false;
            dv.InCellDropdown = true;
            dv.List("\"" + string.Join(",", ConcreteGrades) + "\"");
            dv.InputTitle = "Cấp bền bê tông";
            dv.InputMessage = "Chọn cấp bền bê tông từ danh sách.";

            // Công thức tính fck, fcd từ dropdown
            SetLabelValue(ws, "B5", "C5", "D5", "fck,cube =",
                "=VALUE(SUBSTITUTE(MID(C4,2,99),\",\",\".\"))", "(MPa)", formula: true);
            SetLabelValue(ws, "B6", "C6", "D6", "fck =",
                "=ROUNDDOWN(C5*0.8,0)", "(MPa)", formula: true);
            SetLabelValue(ws, "B7", "C7", "D7", "fcd =",
                "=ROUNDDOWN(G5*C6/G6,0)", "(MPa)", formula: true);

            ws.Range("C5:C7").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        }

        // ── Khối hệ số ───────────────────────────────────────────────────────────
        private static void WriteCoefficients(IXLWorksheet ws, double alphaCc, double gammaC)
        {
            SetBoldUnderline(ws.Cell("F4"), "Hệ số:");

            SetCoeffRow(ws, "F5", "G5", "αcc =", alphaCc);
            SetCoeffRow(ws, "F6", "G6", "γc =", gammaC);

            var formulaRange = ws.Range("F7:G7");
            formulaRange.Merge();
            formulaRange.FirstCell().Value = "fcd = αcc ⋅ fck / γc";
            formulaRange.Style.Font.Italic = true;
        }

        // ── Khối điều kiện ───────────────────────────────────────────────────────
        private static void WriteConditions(IXLWorksheet ws, double colLimit, double wallLimit)
        {
            SetBoldUnderline(ws.Cell("J4"), "Điều kiện:");
            ws.Cell("J5").Value = "ʋd = Ned/(Ac.fcd)";
            ws.Cell("J6").Value = "Cột:";
            ws.Cell("K6").Value = $"ʋd ≤ {colLimit:0.##}";
            ws.Cell("J7").Value = "Vách:";
            ws.Cell("K7").Value = $"ʋd ≤ {wallLimit:0.##}";
        }

        // ── Kết luận tổng ────────────────────────────────────────────────────────
        private static void WriteOverallResult(IXLWorksheet ws)
        {
            SetBoldUnderline(ws.Cell("A8"), "Kết luận:");

            var resultRange = ws.Range("C8:D8");
            resultRange.Merge();
            resultRange.FirstCell().FormulaA1 =
                "=IF(COUNTIF(M11:M2000,\"Không thỏa mãn\")>0,\"Không thỏa mãn\",\"Thỏa mãn\")";
            resultRange.Style.Font.Bold = true;
            resultRange.Style.Font.FontSize = 11;
            resultRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            ApplyConditionalFill(ws.Cell("C8"), "\"Không thỏa mãn\"", "#F8CBAD");
            ApplyConditionalFill(ws.Cell("C8"), "\"Thỏa mãn\"", "#92D050");
        }

        // ── Header bảng ──────────────────────────────────────────────────────────
        private const int HeaderRow = 10;
        private const int FirstDataRow = 11;

        private static readonly string[] Headers =
        {
            "STT", "TẦNG", "LOẠI", "LABEL", "COMBO",
            "Ned\n(kN)", "t3 \n(m)", "t2 \n(m)",
            "Ac\n(m2)", "Ac.fcd\n(kN)", "ʋd", "ʋd\nLimit", "KẾT LUẬN"
        };

        private static void WriteTableHeader(IXLWorksheet ws)
        {
            for (int i = 0; i < Headers.Length; i++)
                ws.Cell(HeaderRow, i + 1).Value = Headers[i];

            var header = ws.Range(HeaderRow, 1, HeaderRow, TotalCols);
            header.Style.Font.Bold = true;
            header.Style.Fill.BackgroundColor = XLColor.FromHtml("#8DB4E2");
            header.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            header.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            header.Style.Alignment.WrapText = true;
            header.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            header.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            ws.Row(HeaderRow).Height = 32;
            ws.SheetView.FreezeRows(HeaderRow);
        }

        // ── Dữ liệu ──────────────────────────────────────────────────────────────
        private static void WriteData(
            IXLWorksheet ws,
            IReadOnlyList<AxialCheckRow> rows,
            double colLimit,
            double wallLimit)
        {
            int r = FirstDataRow;
            foreach (var item in rows)
            {
                ws.Cell(r, ColSTT).Value = item.STT;
                ws.Cell(r, ColTang).Value = item.Story;
                ws.Cell(r, ColLoai).Value = item.ElementType;
                ws.Cell(r, ColLabel).Value = item.Element;
                ws.Cell(r, ColCombo).Value = item.Combo;
                ws.Cell(r, ColNed).Value = Math.Abs(item.Ned);
                ws.Cell(r, ColT3).Value = item.T3;
                ws.Cell(r, ColT2).Value = item.T2;

                // Các cột tính bằng công thức Excel để người dùng có thể chỉnh fcd sau
                ws.Cell(r, ColAc).FormulaA1 = $"=G{r}*H{r}";
                ws.Cell(r, ColAcFcd).FormulaA1 = $"=I{r}*$C$7*1000";
                ws.Cell(r, ColVd).FormulaA1 = $"=ABS(F{r}/J{r})";
                ws.Cell(r, ColVdLimit).Value = item.VdLimit;
                ws.Cell(r, ColKetLuan).FormulaA1 =
                    $"=IF(K{r}<=L{r},\"Thỏa mãn\",\"Không thỏa mãn\")";
                r++;
            }
        }

        // ── Style bảng dữ liệu ───────────────────────────────────────────────────
        private static void ApplyTableStyle(IXLWorksheet ws, int rowCount)
        {
            int lastRow = FirstDataRow + rowCount - 1;

            var table = ws.Range(HeaderRow, 1, lastRow, TotalCols);
            table.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            table.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            table.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            table.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            ws.Range(FirstDataRow, ColNed, lastRow, ColNed).Style.NumberFormat.Format = "0";
            ws.Range(FirstDataRow, ColT3, lastRow, ColT2).Style.NumberFormat.Format = "0.000";
            ws.Range(FirstDataRow, ColAc, lastRow, ColAc).Style.NumberFormat.Format = "0.000";
            ws.Range(FirstDataRow, ColAcFcd, lastRow, ColAcFcd).Style.NumberFormat.Format = "0";
            ws.Range(FirstDataRow, ColVd, lastRow, ColVd).Style.NumberFormat.Format = "0.000";

            var resultRange = ws.Range(FirstDataRow, ColKetLuan, lastRow, ColKetLuan);
            ApplyConditionalFill(resultRange, "\"Không thỏa mãn\"", "#F8CBAD");
            ApplyConditionalFill(resultRange, "\"Thỏa mãn\"", "#E2F0D9");

            int[] colWidths = { 4, 9, 8, 8, 9, 9, 8, 8, 8, 8, 8, 6, 14 };
            for (int i = 0; i < colWidths.Length; i++)
                ws.Column(i + 1).Width = colWidths[i];

            ws.Row(2).Height = 24;
        }

        // ── Helpers ──────────────────────────────────────────────────────────────
        private static void SetBoldUnderline(IXLCell cell, string text)
        {
            cell.Value = text;
            cell.Style.Font.Bold = true;
            cell.Style.Font.Underline = XLFontUnderlineValues.Single;
        }

        private static void SetCoeffRow(IXLWorksheet ws, string labelAddr, string valueAddr, string label, double value)
        {
            ws.Cell(labelAddr).Value = label;
            var cell = ws.Cell(valueAddr);
            cell.Value = value;
            cell.Style.Font.Bold = true;
            cell.Style.Font.FontColor = XLColor.Blue;
        }

        private static void SetLabelValue(
            IXLWorksheet ws,
            string labelAddr, string valueAddr, string unitAddr,
            string label, string valueOrFormula, string unit,
            bool formula = false)
        {
            ws.Cell(labelAddr).Value = label;
            var cell = ws.Cell(valueAddr);
            if (formula) cell.FormulaA1 = valueOrFormula;
            else cell.Value = valueOrFormula;
            ws.Cell(unitAddr).Value = unit;
        }

        private static void ApplyConditionalFill(IXLCell cell, string match, string hexColor)
        {
            cell.AddConditionalFormat()
                .WhenEquals(match)
                .Fill.SetBackgroundColor(XLColor.FromHtml(hexColor));
        }

        private static void ApplyConditionalFill(IXLRange range, string match, string hexColor)
        {
            range.AddConditionalFormat()
                 .WhenEquals(match)
                 .Fill.SetBackgroundColor(XLColor.FromHtml(hexColor));
        }
    }
}
