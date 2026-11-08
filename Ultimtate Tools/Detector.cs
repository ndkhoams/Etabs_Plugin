#nullable enable
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using ClosedXML.Excel;

namespace EtabsLoadAssigner.Core;

/// <summary>
/// Detector: tu dong loc / nhan dien bang tai mong (foundation loads) tu file Excel (.xlsx).
/// Port truc tiep tu etabs_loader/detector.py (app Python) — giu nguyen hanh vi,
/// ke ca cac truong hop bien (nhom Hx/Hy/Vz lech 1-2 cot, dong "note:" ket thuc bang,
/// bang pattern o sheet khac).
///
/// Dang bang nhan duoc (LONG PHU BC-5 A/B):
///     LONG PHU BC-5 A/B |   20 / D    |   20 / E    | ...
///     Foundation loads (action loads) | Hx | Hy | Vz | Hx | Hy | Vz | ...
///     Tolerance +/- 20%                | [kN] ... (dong don vi)
///     Dead loads                |      | 0 | -2 | 19 | 0 | 2 | 19 | ...
///     Wind +Y                   |      | 0 | 7 | -11 | 0 | 6 | -3/+6 | ...
///
/// Va bang phu (thuong o sheet ben canh / cung sheet phia phai):
///     TABLE: Load Pattern Definitions
///     Name | Is Auto Load | Type | Self Weight Multiplier | GUID
///
/// GHI CHU: plugin KHONG doc PDF. File PDF goc (ban ve scan) phai chuyen sang
/// Excel truoc bang script pdf_to_excel/ o repo app Python.
/// </summary>
public static class Detector
{
    // Thanh phan luc (moi diem 3 cot) + mo men (moi diem them 3 cot).
    // Bang hien tai (LONG PHU BC-5) chi co luc — cot mo men TU AN khi file
    // khong co Mx/My/Mz (user chot 08/11 — mo rong cho tuong lai).
    public static readonly string[] ForceComponents = { "Hx", "Hy", "Vz" };
    public static readonly string[] MomentComponents = { "Mx", "My", "Mz" };
    public static readonly string[] Components = { "Hx", "Hy", "Vz" };
    public static readonly string[] AllComponents = { "Hx", "Hy", "Vz", "Mx", "My", "Mz" };

    private static readonly Regex RangeRegex = new(
        @"^\s*([+-]?\d+(?:[.,]\d+)?)\s*/\s*([+-]?)\s*(\d+(?:[.,]\d+)?)\s*$",
        RegexOptions.Compiled);

    private static readonly Regex NumberRegex = new(
        @"^[+-]?\d+(?:[.,]\d+)?$",
        RegexOptions.Compiled);

    internal static string Norm(object? v) => v is null ? "" : v.ToString()!.Trim();

    /// <summary>Parse so kieu Python float(): dau phay thap phan "," -> "." truoc khi parse.</summary>
    internal static double? TryNum(string s)
    {
        if (double.TryParse(s.Replace(',', '.'), NumberStyles.Float,
                CultureInfo.InvariantCulture, out double n))
            return n;
        return null;
    }

    /// <summary>
    /// Port cua parse_cell() trong detector.py.
    /// - null / "" / "-" / "--" / "—" / "n/a" -> blank
    /// - so (ke ca so thuc trong file Excel) -> number
    /// - o viet "+/- 1" / "-/+ 0" / "± 1" (vd bang 90/D.1): 2 gia tri (-V, +V), giu raw goc
    /// - khoang "-3/+6", "67/-71", "-50/48" -> range (gia tri 1, gia tri 2) — KHONG phai min/max
    /// - con lai -> text
    /// Quy tac user chot 08/11: cot "+/-" cua bang KHONG anh huong den gia tri;
    /// chi o tu viet "+/-V" hoac "a/b" moi co 2 gia tri.
    /// </summary>
    public static LoadValue ParseCell(object? raw)
    {
        if (raw is null)
            return new LoadValue { Raw = "", Kind = LoadValueKind.Blank };
        if (raw is double d)
            return new LoadValue { Raw = d.ToString(CultureInfo.InvariantCulture), Kind = LoadValueKind.Number, Value = d };
        if (raw is float f)
            return new LoadValue { Raw = f.ToString(CultureInfo.InvariantCulture), Kind = LoadValueKind.Number, Value = f };
        if (raw is int or long or decimal)
            return new LoadValue { Raw = raw.ToString()!, Kind = LoadValueKind.Number, Value = Convert.ToDouble(raw) };

        string s = Norm(raw);
        if (s is "" or "-" or "--" or "—" or "n/a" or "N/A")
            return new LoadValue { Raw = s, Kind = LoadValueKind.Blank };
        // O viet "+/- 1" / "-/+ 0" / "+ / -10" (ke ca khoang trang): gia tri 1 = -V, gia tri 2 = +V.
        string nospace = s.Replace(" ", "");
        if (nospace.StartsWith("+/-") || nospace.StartsWith("-/+") || nospace.StartsWith("±"))
        {
            string rest = nospace.StartsWith("±") ? nospace.Substring(1) : nospace.Substring(3);
            double? n = TryNum(rest);
            if (n is not null)
                return new LoadValue { Raw = s, Kind = LoadValueKind.Range, VMin = -n.Value, VMax = n.Value };
        }
        Match m = RangeRegex.Match(s);
        if (m.Success)
        {
            double? a = TryNum(m.Groups[1].Value);
            double? b = TryNum(m.Groups[2].Value + m.Groups[3].Value);
            if (a is not null && b is not null)
                return new LoadValue { Raw = s, Kind = LoadValueKind.Range, VMin = a.Value, VMax = b.Value };
        }
        if (NumberRegex.IsMatch(s))
        {
            double? n = TryNum(s);
            if (n is not null)
                return new LoadValue { Raw = s, Kind = LoadValueKind.Number, Value = n.Value };
        }
        return new LoadValue { Raw = s, Kind = LoadValueKind.Text };
    }

    private static bool IsComponentCell(object? v) => Norm(v) is "Hx" or "Hy" or "Vz" or "Mx" or "My" or "Mz";

    /// <summary>
    /// Tim dong tieu de thanh phan. Tra ve (dong, cot Hx dau tien, co mo men).
    /// Co mo men khi nhom dau tien la du 6 cot Hx/Hy/Vz/Mx/My/Mz (moi diem
    /// 6 cot); khong thi moi diem 3 cot luc nhu cu.
    /// </summary>
    private static (int Row, int FirstHxCol, bool HasMoments)? FindComponentHeader(List<List<object?>> grid)
    {
        for (int i = 0; i < grid.Count; i++)
        {
            var row = grid[i];
            var hits = new List<int>();
            for (int j = 0; j < row.Count; j++)
                if (IsComponentCell(row[j]))
                    hits.Add(j);
            if (hits.Count < 3)
                continue;
            var seq = hits.Select(j => Norm(row[j])).ToList();
            if (seq.Contains("Hx") && seq.Contains("Hy") && seq.Contains("Vz"))
            {
                int first = hits[0];
                bool hasMoments = first + 5 < row.Count
                    && Norm(row[first]) == "Hx" && Norm(row[first + 1]) == "Hy" && Norm(row[first + 2]) == "Vz"
                    && Norm(row[first + 3]) == "Mx" && Norm(row[first + 4]) == "My" && Norm(row[first + 5]) == "Mz";
                return (i, first, hasMoments);
            }
        }
        return null;
    }

    private static List<PatternDef> FindPatternTable(List<List<object?>> grid)
    {
        var patterns = new List<PatternDef>();
        for (int i = 0; i < grid.Count; i++)
        {
            var cells = grid[i].Select(Norm).ToList();
            if (!cells.Any(c => c.Contains("Load Pattern Definitions")))
                continue;
            // Dong tieu de cot o 1-2 dong sau.
            for (int j = i + 1; j < Math.Min(i + 4, grid.Count); j++)
            {
                var hdr = grid[j].Select(Norm).ToList();
                if (hdr.Contains("Name") && hdr.Contains("Type"))
                {
                    int nameC = hdr.IndexOf("Name");
                    int typeC = hdr.IndexOf("Type");
                    int? swC = hdr.Contains("Self Weight Multiplier") ? hdr.IndexOf("Self Weight Multiplier") : null;
                    int? guidC = hdr.Contains("GUID") ? hdr.IndexOf("GUID") : null;
                    int? autoC = hdr.Contains("Is Auto Load") ? hdr.IndexOf("Is Auto Load") : null;
                    for (int r = j + 1; r < grid.Count; r++)
                    {
                        var row = grid[r];
                        string nm = nameC < row.Count ? Norm(row[nameC]) : "";
                        if (nm == "")
                            continue;
                        if (nm == "Name")
                            continue;
                        double sw = 0.0;
                        if (swC is not null && swC.Value < row.Count)
                            sw = TryNum(Norm(row[swC.Value])) ?? 0.0;
                        patterns.Add(new PatternDef
                        {
                            Name = nm,
                            IsAuto = autoC is not null && autoC.Value < row.Count ? Norm(row[autoC.Value]) : "",
                            Type = typeC < row.Count ? Norm(row[typeC]) : "",
                            SelfWeight = sw,
                            Guid = guidC is not null && guidC.Value < row.Count ? Norm(row[guidC.Value]) : "",
                        });
                    }
                    return patterns;
                }
            }
            return patterns;
        }
        return patterns;
    }

    /// <summary>
    /// Nhận diện bảng tai mong tu mot grid (list[list]) — dung chung cho Excel.
    /// Port truc tiep detect_from_grid() trong detector.py.
    /// </summary>
    public static FoundationTable DetectFromGrid(List<List<object?>> grid, string source = "", string sheet = "")
    {
        var table = new FoundationTable { Source = source, Sheet = sheet };
        table.Patterns = FindPatternTable(grid);
        var found = FindComponentHeader(grid);
        if (found is null)
        {
            table.Warnings.Add(
                "Không tìm thấy dòng tiêu đề Hx/Hy/Vz trong file. " +
                "Kiểm tra lại: bảng cần có các cột Hx, Hy, Vz (mỗi điểm 3 cột).");
            return table;
        }
        int hdrRow = found.Value.Row;
        int firstHxCol = found.Value.FirstHxCol;
        bool hasMoments = found.Value.HasMoments;
        table.HasMoments = hasMoments;
        // So cot moi diem: 3 (chi luc) hoac 6 (luc + mo men).
        string[] comps = hasMoments ? AllComponents : ForceComponents;
        int groupSize = comps.Length;

        // --- Dong ten diem: dong ngay tren dong Hx/Hy/Vz, doc ten o cot Hx ---
        var points = new List<string>();
        var pointCols = new List<int>();
        int j = firstHxCol;
        var pointRow = hdrRow >= 1 ? grid[hdrRow - 1] : new List<object?>();
        var rowVals = grid[hdrRow];
        while (j < rowVals.Count)
        {
            string[] Group(int jj)
            {
                var g = new string[groupSize];
                for (int k = 0; k < groupSize; k++)
                    g[k] = jj + k < rowVals.Count ? Norm(rowVals[jj + k]) : "";
                return g;
            }
            bool MatchesGroup(string[] g)
            {
                for (int k = 0; k < groupSize; k++)
                    if (g[k] != comps[k])
                        return false;
                return true;
            }
            string[] group = Group(j);
            if (!MatchesGroup(group))
            {
                // Nhom lech (co the co cot trong) — dich sang phai 1 cot, thu toi da 2 lan.
                bool shifted = false;
                foreach (int d in new[] { 1, 2 })
                {
                    string[] group2 = Group(j + d);
                    if (MatchesGroup(group2))
                    {
                        j += d;
                        group = group2;
                        shifted = true;
                        break;
                    }
                }
                if (!shifted)
                    break;
            }
            string name = j < pointRow.Count ? Norm(pointRow[j]) : "";
            if (name == "")
            {
                name = $"Point {points.Count + 1}";
                table.Warnings.Add($"Không đọc được tên điểm ở cột {j + 1}, tạm đặt '{name}'.");
            }
            points.Add(name);
            pointCols.Add(j);
            j += groupSize;
        }
        if (points.Count == 0)
        {
            table.Warnings.Add("Tìm thấy dòng Hx/Hy/Vz nhưng không tách được điểm nào.");
            return table;
        }
        table.Points = points;

        // --- Tieu de bang: cac dong tren va ngay dong Hx (cot A thuong chua tieu de) ---
        for (int r = Math.Max(0, hdrRow - 3); r <= hdrRow; r++)
            foreach (object? c in grid[r])
                if (Norm(c).Contains("Foundation loads"))
                    table.Title = Norm(c);

        // --- Dong du lieu: bat dau sau dong Hx/Hy/Vz, bo dong don vi [kN] ---
        int i = hdrRow + 1;
        if (i < grid.Count && grid[i].Any(c => Norm(c).Contains("[kN]")))
            i += 1;
        while (i < grid.Count)
        {
            var row = grid[i];
            i += 1;
            if (row.Count == 0)
                continue;
            string first = Norm(row[0]);
            if (first == "")
                continue;
            string low = first.ToLowerInvariant();
            if (low.StartsWith("note:") || (low.StartsWith("note") && low.Contains("applied loads")))
                break;
            if (low.Contains("foundation loads") || low.Contains("long phu"))
                break; // Bang moi (vd diem 27/D o duoi trong cung sheet).
            if (first is "Hx" or "Hy" or "Vz" or "Mx" or "My" or "Mz")
                continue;
            // Cot +/-: thuong la cot thu 2 (index 1), hoac ngay sau ten tai.
            bool reversible = row.Take(3).Any(c => Norm(c) is "+/-" or "±" or "+/- ");
            var vals = new Dictionary<string, Dictionary<string, LoadValue>>();
            for (int p = 0; p < points.Count; p++)
            {
                int col = pointCols[p];
                var comp = new Dictionary<string, LoadValue>();
                for (int k = 0; k < comps.Length; k++)
                {
                    object? cell = col + k < row.Count ? row[col + k] : null;
                    comp[comps[k]] = ParseCell(cell);
                }
                vals[points[p]] = comp;
            }
            table.Rows.Add(new FoundationRow { Name = first, Reversible = reversible, Values = vals });
        }
        if (table.Rows.Count == 0)
            table.Warnings.Add("Nhận được điểm nhưng không đọc được dòng tải nào.");
        return table;
    }

    private static object? CellValue(IXLCell cell)
    {
        var v = cell.Value;
        if (v.IsBlank)
            return null;
        if (v.IsNumber)
            return v.GetNumber();
        if (v.IsText)
            return v.GetText();
        return v.ToString();
    }

    /// <summary>
    /// Doc toan bo sheet thanh grid (list[list]) — dung chung cho Excel.
    /// TOI UU (2026-11-08): truoc day duyet TUNG o bang ws.Cell(r, c) tren
    /// toan bo LastRowUsed x LastColumnUsed — ClosedXML tao object cell cho
    /// ca o trong (rat cham voi sheet A0 thua). Gio chi liet ke cac o DA
    /// DUNG 1 lan (ws.CellsUsed()) roi dien vao grid da cap phat san.
    /// Grid tra ve GIONG HET ban cu (cung so dong/cot, o trong = null) nen
    /// DetectFromGrid/FindPatternTable cho cung ket qua.
    /// </summary>
    private static List<List<object?>> ReadGrid(IXLWorksheet ws)
    {
        var grid = new List<List<object?>>();
        var lastRow = ws.LastRowUsed();
        if (lastRow is null)
            return grid;
        int nRows = lastRow.RowNumber();
        int lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;
        for (int r = 0; r < nRows; r++)
        {
            var row = new List<object?>(lastCol);
            for (int c = 0; c < lastCol; c++)
                row.Add(null);
            grid.Add(row);
        }
        foreach (var cell in ws.CellsUsed())
        {
            int r = cell.Address.RowNumber - 1;
            int c = cell.Address.ColumnNumber - 1;
            if (r >= 0 && r < nRows && c >= 0 && c < lastCol)
                grid[r][c] = CellValue(cell);
        }
        return grid;
    }

    /// <summary>
    /// Doc grid chi cho cac dong [firstRow..lastRow] (1-based, bao gom ca 2
    /// dau) — dung khi chi can doc phan cuoi sheet (vd bang pattern nam
    /// duoi bang tai), khong can doc ca sheet.
    /// </summary>
    private static List<List<object?>> ReadGridRange(IXLWorksheet ws, int firstRow, int lastRow)
    {
        var grid = new List<List<object?>>();
        int lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;
        int nRows = Math.Max(0, lastRow - firstRow + 1);
        for (int i = 0; i < nRows; i++)
        {
            var row = new List<object?>(lastCol);
            for (int c = 0; c < lastCol; c++)
                row.Add(null);
            grid.Add(row);
        }
        foreach (var cell in ws.CellsUsed())
        {
            int r = cell.Address.RowNumber - firstRow;
            int c = cell.Address.ColumnNumber - 1;
            if (r >= 0 && r < nRows && c >= 0 && c < lastCol)
                grid[r][c] = CellValue(cell);
        }
        return grid;
    }

    /// <summary>
    /// Tim bang pattern tren 1 sheet MA KHONG build full grid: chi quet cac
    /// o da dung de tim dong chua "Load Pattern Definitions", roi doc grid
    /// chi tu dong do tro xuong va dung lai FindPatternTable(grid) cu.
    /// Ket qua giong het FindPatternTable(ReadGrid(ws)) vi dong tieu de la
    /// dong 0 cua sub-grid, moi vi tri tuong doi giu nguyen.
    /// </summary>
    private static List<PatternDef> FindPatternTableOnSheet(IXLWorksheet ws)
    {
        int? titleRow = null;
        foreach (var cell in ws.CellsUsed())
        {
            if (Norm(CellValue(cell)).Contains("Load Pattern Definitions"))
            {
                int r = cell.Address.RowNumber;
                if (titleRow is null || r < titleRow.Value)
                    titleRow = r;
            }
        }
        if (titleRow is null)
            return new List<PatternDef>();
        int lastRow = ws.LastRowUsed()?.RowNumber() ?? titleRow.Value;
        return FindPatternTable(ReadGridRange(ws, titleRow.Value, lastRow));
    }

    /// <summary>Liet ke ten sheet trong file Excel (de user chon bang can nhap khi file co nhieu bang).</summary>
    public static List<string> ListSheets(string path)
    {
        // TOI UU (2026-11-08): doc ten sheet truc tiep tu xl/workbook.xml
        // trong file zip (khong can mo toan bo workbook bang ClosedXML) —
        // nhanh hon ~40 lan. Loi thi roi ve ClosedXML (hanh vi cu).
        try
        {
            var names = ReadSheetNamesFromZip(path);
            if (names is not null && names.Count > 0)
                return names;
        }
        catch { /* fallback xuong duoi */ }
        using var wb = new XLWorkbook(path);
        return wb.Worksheets.Select(ws => ws.Name).ToList();
    }

    /// <summary>Ten sheet luu thong tin match trong file Excel noi luc (user chot 08/11 — thay file .json rieng).</summary>
    public const string ProfileSheetName = "LoadAssigner";

    /// <summary>
    /// Doc thong tin match tu sheet "LoadAssigner" trong file Excel (neu co).
    /// Tra null neu khong co sheet. Doc kieu chiu dung: quet tat ca dong, phan
    /// loai theo cot A (ten tai / Hx-Hy-Vz / "Ghi de tai cu" / tieu de bo qua).
    /// </summary>
    public static Profile? ReadProfileSheet(string path)
    {
        using var wb = new XLWorkbook(path);
        var ws = wb.Worksheets.FirstOrDefault(w => string.Equals(w.Name, ProfileSheetName, StringComparison.OrdinalIgnoreCase));
        if (ws is null)
            return null;
        var p = new Profile();
        foreach (var row in ws.RowsUsed())
        {
            string a = (row.Cell(1).GetString() ?? "").Trim();
            if (a == "")
                continue;
            if (a == "LoadAssignerProfile" || a == "Tải" || a == "Phuong & chieu")
                continue;
            string b = (row.Cell(2).GetString() ?? "").Trim();
            string c = (row.Cell(3).GetString() ?? "").Trim();
            if (a == "Hx" || a == "Hy" || a == "Vz")
            {
                string axis = b == "" ? (a == "Hx" ? "X" : a == "Hy" ? "Y" : "Z") : b;
                int sign = c == "-" ? -1 : +1;
                switch (a)
                {
                    case "Hx": p.Axis.HxAxis = axis; p.Axis.HxSign = sign; break;
                    case "Hy": p.Axis.HyAxis = axis; p.Axis.HySign = sign; break;
                    default: p.Axis.VzAxis = axis; p.Axis.VzSign = sign; break;
                }
                continue;
            }
            if (a == "Ghi de tai cu")
            {
                p.ReplaceExisting = b.Equals("TRUE", StringComparison.OrdinalIgnoreCase)
                    || b == "1" || b.Equals("Có", StringComparison.OrdinalIgnoreCase)
                    || b.Equals("Yes", StringComparison.OrdinalIgnoreCase);
                continue;
            }
            p.PatternMap[a] = new RowPattern { Left = b, Right = c };
        }
        return p;
    }

    /// <summary>
    /// Ghi thong tin match vao sheet "LoadAssigner" trong file Excel (xoa sheet
    /// cu roi tao moi — dam bao noi dung sach). Ghi de file goc; file dang mo
    /// trong Excel thi se nem loi (file bi khoa) — caller bat va bao user dong
    /// file truoc. Cac sheet khac giu nguyen. TU CHOI ghi neu bang rong
    /// (0 dong tai) — tranh ghi de sheet thanh trang nhu loi truoc day.
    /// </summary>
    public static void WriteProfileSheet(string path, Profile profile, FoundationTable table)
    {
        if (table.Rows.Count == 0)
            throw new InvalidOperationException("Bảng tải chưa có dòng nào — chọn đúng sheet có bảng tải trước khi lưu.");
        using var wb = new XLWorkbook(path);
        var old = wb.Worksheets.FirstOrDefault(w => string.Equals(w.Name, ProfileSheetName, StringComparison.OrdinalIgnoreCase));
        old?.Delete();
        var ws = wb.AddWorksheet(ProfileSheetName);
        ws.Cell(1, 1).Value = "LoadAssignerProfile";
        ws.Cell(1, 2).Value = "Match pattern & cai dat — plugin tu ghi, khong sua tay neu khong ro.";
        ws.Cell(3, 1).Value = "Tải";
        ws.Cell(3, 2).Value = "Pattern — giá trị 1 (trái)";
        ws.Cell(3, 3).Value = "Pattern — giá trị 2 (phải)";
        int r = 4;
        foreach (var row in table.Rows)
        {
            var rp = profile.PatternFor(row.Name);
            ws.Cell(r, 1).Value = row.Name;
            ws.Cell(r, 2).Value = rp.Left;
            ws.Cell(r, 3).Value = rp.Right;
            r++;
        }
        r++; // dong trong
        ws.Cell(r, 1).Value = "Phuong & chieu";
        r++;
        ws.Cell(r, 1).Value = "Hx"; ws.Cell(r, 2).Value = profile.Axis.HxAxis; ws.Cell(r, 3).Value = profile.Axis.HxSign < 0 ? "-" : "+";
        r++;
        ws.Cell(r, 1).Value = "Hy"; ws.Cell(r, 2).Value = profile.Axis.HyAxis; ws.Cell(r, 3).Value = profile.Axis.HySign < 0 ? "-" : "+";
        r++;
        ws.Cell(r, 1).Value = "Vz"; ws.Cell(r, 2).Value = profile.Axis.VzAxis; ws.Cell(r, 3).Value = profile.Axis.VzSign < 0 ? "-" : "+";
        r++;
        ws.Cell(r, 1).Value = "Ghi de tai cu"; ws.Cell(r, 2).Value = profile.ReplaceExisting ? "TRUE" : "FALSE";
        ws.Column(1).Width = 45;
        ws.Column(2).Width = 40;
        ws.Column(3).Width = 40;
        wb.Save();
    }

    /// <summary>Doc ten sheet tu xl/workbook.xml (zip). Tra null neu khong doc duoc.</summary>
    private static List<string>? ReadSheetNamesFromZip(string path)
    {
        using var fs = File.OpenRead(path);
        using var zip = new ZipArchive(fs, ZipArchiveMode.Read);
        var entry = zip.GetEntry("xl/workbook.xml");
        if (entry is null)
            return null;
        using var s = entry.Open();
        using var reader = XmlReader.Create(s);
        var names = new List<string>();
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "sheet")
            {
                string? nm = reader.GetAttribute("name");
                if (nm is not null)
                    names.Add(nm);
            }
        }
        return names.Count > 0 ? names : null;
    }

    public static FoundationTable DetectExcel(string path, string? sheet = null)
    {
        using var wb = new XLWorkbook(path);
        FoundationTable? best = null;
        if (sheet is not null)
        {
            // TOI UU (2026-11-08): truoc day vong 2 doc lai TOAN BO grid cua
            // moi sheet khac chi de gom pattern table. Gio moi sheet khac chi
            // quet nhe de tim bang pattern (FindPatternTableOnSheet), khong
            // build full grid. Workbook chi mo 1 lan.
            var target = wb.Worksheets.FirstOrDefault(w => w.Name == sheet);
            if (target is null)
                return new FoundationTable
                {
                    Source = path,
                    Warnings = { "File Excel không có sheet nào." },
                };
            best = DetectFromGrid(ReadGrid(target), path, target.Name);
            foreach (var ws in wb.Worksheets)
            {
                if (ws.Name == sheet)
                    continue;
                best.Patterns = best.Patterns.Concat(FindPatternTableOnSheet(ws)).ToList();
            }
            return best;
        }
        foreach (var ws in wb.Worksheets)
        {
            var t = DetectFromGrid(ReadGrid(ws), path, ws.Name);
            // Uu tien bang co ca diem + dong tai; gop pattern table tu moi sheet.
            if (best is null)
                best = t;
            else if (t.Rows.Count > best.Rows.Count)
            {
                t.Patterns = best.Patterns.Concat(t.Patterns).ToList();
                best = t;
            }
            else
            {
                best.Patterns = best.Patterns.Concat(t.Patterns).ToList();
            }
        }
        if (best is null)
            return new FoundationTable
            {
                Source = path,
                Warnings = { "File Excel không có sheet nào." },
            };
        return best;
    }

    public static FoundationTable DetectFile(string path, string? sheet = null)
    {
        string low = path.ToLowerInvariant();
        if (low.EndsWith(".xlsx") || low.EndsWith(".xlsm"))
            return DetectExcel(path, sheet);
        if (low.EndsWith(".pdf"))
            throw new NotSupportedException(
                "Plugin chỉ đọc file Excel (.xlsx). File PDF bản vẽ là ảnh scan — " +
                "chuyển sang Excel trước bằng script pdf_to_excel/ ở repo app Python.");
        if (low.EndsWith(".xls"))
            throw new NotSupportedException("File .xls cũ không đọc được — lưu lại thành .xlsx rồi import.");
        throw new NotSupportedException($"Không hỗ trợ định dạng file: {path} (chỉ .xlsx).");
    }
}

/// <summary>Mot o gia tri trong bang: so, khoang 'a/b' (VMin = gia tri 1, VMax = gia tri 2 — KHONG phai min/max), hoac trong.</summary>
public sealed class LoadValue
{
    public string Raw { get; set; } = "";
    public LoadValueKind Kind { get; set; } = LoadValueKind.Blank;
    public double? Value { get; set; }
    public double? VMin { get; set; }
    public double? VMax { get; set; }

    public bool IsZero() => Kind == LoadValueKind.Number && Value == 0;
}

public enum LoadValueKind
{
    Blank,
    Number,
    Range,
    Text,
}

public sealed class FoundationRow
{
    public string Name { get; set; } = "";
    // Cot "+/-" cua bang — CHI de hien thi "(+/-)" o grid/match (user tu xu ly trong combo sau).
    // KHONG anh huong den gia tri gan (user chot 08/11).
    public bool Reversible { get; set; }
    public Dictionary<string, Dictionary<string, LoadValue>> Values { get; set; } = new();

    /// <summary>Co o 2 gia tri (khoang "a/b" hoac "+/-V") o bat ky diem/thanh phan nao — dung de goi y pattern cot phai.</summary>
    public bool HasRangeCell()
    {
        foreach (var compVals in Values.Values)
            foreach (var lv in compVals.Values)
                if (lv.Kind == LoadValueKind.Range)
                    return true;
        return false;
    }

    public bool IsAllZero()
    {
        foreach (var compVals in Values.Values)
            foreach (var lv in compVals.Values)
            {
                if (lv.Kind == LoadValueKind.Number && lv.Value != 0)
                    return false;
                if (lv.Kind == LoadValueKind.Range && (lv.VMin != 0 || lv.VMax != 0))
                    return false;
            }
        return true;
    }
}

public sealed class PatternDef
{
    public string Name { get; set; } = "";
    public string IsAuto { get; set; } = "";
    public string Type { get; set; } = "";
    public double SelfWeight { get; set; }
    public string Guid { get; set; } = "";
}

public sealed class FoundationTable
{
    public string Title { get; set; } = "";
    public List<string> Points { get; set; } = new(); // vd ["20 / D", "20 / E"]
    public List<FoundationRow> Rows { get; set; } = new();
    public List<PatternDef> Patterns { get; set; } = new();
    public string Source { get; set; } = "";
    public string Sheet { get; set; } = "";
    public List<string> Warnings { get; set; } = new();
    // True khi bang co cot mo men Mx/My/Mz (moi diem 6 cot). False = chi co
    // luc Hx/Hy/Vz — cot mo men o grid TU AN (user chot 08/11).
    public bool HasMoments { get; set; }

    public Dictionary<string, LoadValue>? PointValues(string point, string rowName)
    {
        foreach (var r in Rows)
            if (r.Name == rowName)
                return r.Values.TryGetValue(point, out var v) ? v : null;
        return null;
    }
}
