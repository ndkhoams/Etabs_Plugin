#nullable enable
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
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
    public static readonly string[] Components = { "Hx", "Hy", "Vz" };

    private static readonly Regex RangeRegex = new(
        @"^\s*([+-]?\d+(?:[.,]\d+)?)\s*/\s*([+-])\s*(\d+(?:[.,]\d+)?)\s*$",
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
    /// - o viet "+/- 1" / "-/+ 0" / "± 1" (vd bang 90/D.1): lay do lon, giu raw goc
    /// - khoang "-3/+6", "67/-71" -> range (vmin, vmax)
    /// - con lai -> text
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
        // O viet "+/- 1" / "-/+ 0": gia tri dao chieu, lay do lon; dau +/- goc giu trong raw.
        if (s.StartsWith("+/-") || s.StartsWith("-/+") || s.StartsWith("±"))
        {
            string rest = (s.StartsWith("+/-") || s.StartsWith("-/+")) ? s.Substring(3).Trim() : s.Substring(1).Trim();
            double? n = TryNum(rest);
            if (n is not null)
                return new LoadValue { Raw = s, Kind = LoadValueKind.Number, Value = n.Value };
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

    private static bool IsComponentCell(object? v) => Norm(v) is "Hx" or "Hy" or "Vz";

    private static (int Row, int FirstHxCol)? FindComponentHeader(List<List<object?>> grid)
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
                return (i, hits[0]);
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

        // --- Dong ten diem: dong ngay tren dong Hx/Hy/Vz, doc ten o cot Hx ---
        var points = new List<string>();
        var pointCols = new List<int>();
        int j = firstHxCol;
        var pointRow = hdrRow >= 1 ? grid[hdrRow - 1] : new List<object?>();
        var rowVals = grid[hdrRow];
        while (j < rowVals.Count)
        {
            string[] Trio(int jj) => new[]
            {
                jj < rowVals.Count ? Norm(rowVals[jj]) : "",
                jj + 1 < rowVals.Count ? Norm(rowVals[jj + 1]) : "",
                jj + 2 < rowVals.Count ? Norm(rowVals[jj + 2]) : "",
            };
            string[] trio = Trio(j);
            if (!(trio[0] == "Hx" && trio[1] == "Hy" && trio[2] == "Vz"))
            {
                // Nhom lech (co the co cot trong) — dich sang phai 1 cot, thu toi da 2 lan.
                bool shifted = false;
                foreach (int d in new[] { 1, 2 })
                {
                    string[] trio2 = Trio(j + d);
                    if (trio2[0] == "Hx" && trio2[1] == "Hy" && trio2[2] == "Vz")
                    {
                        j += d;
                        trio = trio2;
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
            j += 3;
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
            if (first is "Hx" or "Hy" or "Vz")
                continue;
            // Cot +/-: thuong la cot thu 2 (index 1), hoac ngay sau ten tai.
            bool reversible = row.Take(3).Any(c => Norm(c) is "+/-" or "±" or "+/- ");
            var vals = new Dictionary<string, Dictionary<string, LoadValue>>();
            for (int p = 0; p < points.Count; p++)
            {
                int col = pointCols[p];
                var comp = new Dictionary<string, LoadValue>();
                for (int k = 0; k < Components.Length; k++)
                {
                    object? cell = col + k < row.Count ? row[col + k] : null;
                    comp[Components[k]] = ParseCell(cell);
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

    private static List<List<object?>> ReadGrid(IXLWorksheet ws)
    {
        var grid = new List<List<object?>>();
        var lastRow = ws.LastRowUsed();
        if (lastRow is null)
            return grid;
        int lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;
        for (int r = 1; r <= lastRow.RowNumber(); r++)
        {
            var row = new List<object?>();
            for (int c = 1; c <= lastCol; c++)
                row.Add(CellValue(ws.Cell(r, c)));
            grid.Add(row);
        }
        return grid;
    }

    /// <summary>Liet ke ten sheet trong file Excel (de user chon bang can nhap khi file co nhieu bang).</summary>
    public static List<string> ListSheets(string path)
    {
        using var wb = new XLWorkbook(path);
        return wb.Worksheets.Select(ws => ws.Name).ToList();
    }

    public static FoundationTable DetectExcel(string path, string? sheet = null)
    {
        using var wb = new XLWorkbook(path);
        FoundationTable? best = null;
        foreach (var ws in wb.Worksheets)
        {
            if (sheet is not null && ws.Name != sheet)
                continue;
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
        // Neu chi dinh sheet: van gop bang pattern tu cac sheet con lai.
        if (sheet is not null)
        {
            foreach (var ws in wb.Worksheets)
            {
                if (ws.Name == sheet)
                    continue;
                best.Patterns = best.Patterns.Concat(FindPatternTable(ReadGrid(ws))).ToList();
            }
        }
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

/// <summary>Mot o gia tri trong bang: so, khoang '-3/+6', hoac trong.</summary>
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
    public bool Reversible { get; set; } // Cot "+/-" — tai co the dao chieu.
    public Dictionary<string, Dictionary<string, LoadValue>> Values { get; set; } = new();

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

    public Dictionary<string, LoadValue>? PointValues(string point, string rowName)
    {
        foreach (var r in Rows)
            if (r.Name == rowName)
                return r.Values.TryGetValue(point, out var v) ? v : null;
        return null;
    }
}
