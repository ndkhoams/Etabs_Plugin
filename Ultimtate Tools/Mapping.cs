#nullable enable
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace EtabsLoadAssigner.Core;

/// <summary>
/// Mapping: phuong/chieu Hx-Hy-Vz -> ETABS global, match pattern,
/// va profile luu/tai cai dat. Port truc tiep tu etabs_loader/mapping.py.
///
/// Quy uoc mac dinh (suy tu so do toa do trong anh user gui):
///   - So do ve: Hx mui ten sang trai, Hy len tren, "+z down".
///   - ETABS global: +Z huong len. Nen Vz (duong khi huong xuong) -> Fz = -Vz.
///   - Hx/Hy: dau + mac dinh giu nguyen (+1). Dau nay KHONG suy ra duoc tu anh —
///     phai kiem tra 1 diem that roi dao dau trong bang neu can.
/// </summary>
public sealed class AxisMapping
{
    public string HxAxis { get; set; } = "X";
    public int HxSign { get; set; } = +1;
    public string HyAxis { get; set; } = "Y";
    public int HySign { get; set; } = +1;
    public string VzAxis { get; set; } = "Z";
    public int VzSign { get; set; } = -1; // +z down (client) -> ETABS +Z up

    public List<string> Validate()
    {
        var errs = new List<string>();
        var axes = new[] { HxAxis, HyAxis, VzAxis }.OrderBy(a => a).ToArray();
        if (!axes.SequenceEqual(new[] { "X", "Y", "Z" }))
            errs.Add("Hx/Hy/Vz phải ánh xạ sang 3 trục X, Y, Z khác nhau (không trùng trục).");
        foreach (int s in new[] { HxSign, HySign, VzSign })
            if (s is not (+1 or -1))
                errs.Add($"Dấu phải +1 hoặc -1, nhận được {s}.");
        return errs;
    }

    /// <summary>Tra ve (Fx, Fy, Fz) theo truc ETABS.</summary>
    public (double Fx, double Fy, double Fz) Apply(double hx, double hy, double vz)
    {
        var outMap = new Dictionary<string, double> { ["X"] = 0, ["Y"] = 0, ["Z"] = 0 };
        outMap[HxAxis] = HxSign * hx;
        outMap[HyAxis] = HySign * hy;
        outMap[VzAxis] = VzSign * vz;
        return (outMap["X"], outMap["Y"], outMap["Z"]);
    }

    public string Describe()
    {
        static string Fmt(string axis, int sign) => $"{(sign > 0 ? "+" : "-")}{axis}";
        return $"Hx -> F{Fmt(HxAxis, HxSign)} | Hy -> F{Fmt(HyAxis, HySign)} | Vz -> F{Fmt(VzAxis, VzSign)}";
    }

    public Dictionary<string, object?> ToDictionary() => new()
    {
        ["hx_axis"] = HxAxis,
        ["hx_sign"] = HxSign,
        ["hy_axis"] = HyAxis,
        ["hy_sign"] = HySign,
        ["vz_axis"] = VzAxis,
        ["vz_sign"] = VzSign,
    };

    public static AxisMapping FromDictionary(JsonElement d)
    {
        if (d.ValueKind != JsonValueKind.Object)
            return new AxisMapping();
        return new AxisMapping
        {
            HxAxis = JsonHelpers.GetString(d, "hx_axis", "X"),
            HxSign = JsonHelpers.GetInt(d, "hx_sign", +1),
            HyAxis = JsonHelpers.GetString(d, "hy_axis", "Y"),
            HySign = JsonHelpers.GetInt(d, "hy_sign", +1),
            VzAxis = JsonHelpers.GetString(d, "vz_axis", "Z"),
            VzSign = JsonHelpers.GetInt(d, "vz_sign", -1),
        };
    }
}

/// <summary>Match pattern 2 cot + phuong/chieu. Port cua mapping.py (da bo ResolveRange — khoang gio tach min/max sang 2 cot).</summary>
public static class Mapping
{
/// <summary>
/// Goi y pattern cot PHAI (gia tri lon) cho 1 dong co +/-: doi xung cua
/// pattern cot trai. Thu theo thu tu: "<trai>_NEG" (neu co trong model),
/// doi dau "+" <-> "-" trong ten (WX+ -> WX-, WY+1 -> WY-1). Khong co thi
/// tra "" — user tu chon. Tool KHONG tu tao pattern (user chot 08/11).
/// </summary>
public static string SuggestPatternOpposite(string leftPattern, IReadOnlyList<string> available)
{
    if (leftPattern == "" || available.Count == 0)
        return "";
    string neg = leftPattern + "_NEG";
    if (available.Contains(neg))
        return neg;
    string swapped = leftPattern.Contains('+') ? leftPattern.Replace('+', '-')
        : leftPattern.Contains('-') ? leftPattern.Replace('-', '+') : "";
    if (swapped != "" && available.Contains(swapped))
        return swapped;
    return "";
}

    /// <summary>Dinh dang so kieu Python str(float): 6.0 chu khong phai "6" (dung trong ghi chu).</summary>
    internal static string PyNum(double v) =>
        v.ToString("0.0###############", CultureInfo.InvariantCulture);

    // --- Goi y match pattern (CHI LA GOI Y, user chot trong bang) ---------
    // Ten pattern ETABS lay tu bang 'Load Pattern Definitions' trong anh:
    // BC1, BC2, BT1, BT2, DLS, EX, EY, EZ, F, LL1..LL5, LLR, NX, NY,
    // SLL1, SLL2, SPL1, SPL2, SW, T, WX-, WX+, WY-1, WY-2, WY+1, WY+2.
    // Cac goi y duoi day suy tu ten goi (BC=belt conveyor, BT=belt tension,
    // NX/NY=notional/imperfection, F=friction...) — chua duoc user xac nhan.
    public static readonly IReadOnlyDictionary<string, string> SuggestedMap =
        new Dictionary<string, string>
        {
            ["dead loads"] = "DLS",
            ["live load belt conveyor load bc-5a"] = "BC1",
            ["live load belt conveyor load bc-5b"] = "BC2",
            ["live load belt tension bc-5a max/min"] = "BT1",
            ["live load belt tension bc-5b max/min"] = "BT2",
            ["live load roof (inacc.)"] = "LLR",
            ["imperfektion x"] = "NX",
            ["imperfektion y"] = "NY",
            ["friction +x"] = "F",
            ["wind +x"] = "WX+",
            ["wind -x"] = "WX-",
            ["wind +y"] = "WY+1",
            ["wind -y"] = "WY-1",
            ["temperatur dt = 16°c"] = "T",
            ["earthquake x (lfp; 0,12g)"] = "EX",
            ["earthquake y (lfp; 0,12g)"] = "EY",
            ["earthquake z (lfp; 0,94g)"] = "EZ",
            ["live load walkway left side (5,0kn/m²)"] = "LL1",
            ["live load walkway middle (5,0kn/m²)"] = "LL2",
            ["live load walkway right side (5,0kn/m²)"] = "LL3",
            ["live load platforms and stairs"] = "LL4",
            ["live load lifting jack #1 (service load 16 kn)"] = "LL5",
        };

    public static string NormName(string s) =>
        Regex.Replace(s.ToLowerInvariant().Trim(), @"\s+", " ");

    /// <summary>Goi y pattern ETABS cho 1 dong tai Excel. Tra ve "" neu khong chac.</summary>
    public static string SuggestPattern(string excelName, IReadOnlyList<string> available)
    {
        if (available.Count == 0)
            return "";
        string key = NormName(excelName);
        if (SuggestedMap.TryGetValue(key, out string? mapped) && available.Contains(mapped))
            return mapped;
        // Cham diem theo token chung (vd 'wind +x' trung 'wx+').
        string best = "";
        int bestScore = 0;
        var tokens = new HashSet<string>(Regex.Matches(key, "[a-z0-9]+").Cast<Match>().Select(m => m.Value));
        foreach (string pat in available)
        {
            var ptokens = new HashSet<string>(Regex.Matches(pat.ToLowerInvariant(), "[a-z0-9]+").Cast<Match>().Select(m => m.Value));
            int score = tokens.Intersect(ptokens).Count();
            if (key.Contains(pat.ToLowerInvariant()))
                score += 2;
            if (score > bestScore)
            {
                best = pat;
                bestScore = score;
            }
        }
        return bestScore >= 1 ? best : "";
    }
}

/// <summary>2 pattern match cho 1 dong tai: cot trai (gia tri nho / gia tri don), cot phai (gia tri lon, chi khi co +/-).</summary>
public sealed class RowPattern
{
    public string Left { get; set; } = "";
    public string Right { get; set; } = "";
}

/// <summary>Cai dat luu lai de gan diem moi khong phai lam lai. Port cua Profile trong mapping.py.</summary>
public sealed class Profile
{
    public int Version { get; set; } = 1;
    public string SourceFile { get; set; } = "";
    public string Sheet { get; set; } = "";
    public AxisMapping Axis { get; set; } = new();
    /// <summary>
    /// Match 2 cot (user chot 08/11): cot TRAI = pattern nhan gia tri NHO
    /// (gia tri don, hoac min cua o khoang / -V cua hang +/-); cot PHAI =
    /// pattern nhan gia tri LON (max cua o khoang / +V cua hang +/-). Dong
    /// khong co dau +/- chi can cot trai. User tu dinh nghia du pattern
    /// trong ETABS — tool KHONG tu tao pattern.
    /// </summary>
    public Dictionary<string, RowPattern> PatternMap { get; set; } = new();
    public bool ReplaceExisting { get; set; } = true;
    public bool SkipZeroRows { get; set; } = true;

    /// <summary>
    /// 2 pattern ETABS cho 1 dong tai: khop dung ten truoc, roi khop khong
    /// phan biet hoa/thuong. Tra ve RowPattern rong neu chua match.
    /// </summary>
    public RowPattern PatternFor(string rowName)
    {
        if (PatternMap.TryGetValue(rowName, out RowPattern? exact))
            return exact;
        string key = Mapping.NormName(rowName);
        foreach (var kv in PatternMap)
            if (Mapping.NormName(kv.Key) == key)
                return kv.Value;
        return new RowPattern();
    }

    public Dictionary<string, object?> ToDictionary() => new()
    {
        ["version"] = Version,
        ["source_file"] = SourceFile,
        ["sheet"] = Sheet,
        ["axis"] = Axis.ToDictionary(),
        ["pattern_map"] = PatternMap.ToDictionary(kv => kv.Key,
            kv => (object?)new Dictionary<string, object?>
            {
                ["left"] = kv.Value.Left,
                ["right"] = kv.Value.Right,
            }),
        ["replace_existing"] = ReplaceExisting,
        ["skip_zero_rows"] = SkipZeroRows,
    };

    public static Profile FromDictionary(JsonElement d)
    {
        var p = new Profile();
        if (d.ValueKind != JsonValueKind.Object)
            return p;
        p.Version = JsonHelpers.GetInt(d, "version", 1);
        p.SourceFile = JsonHelpers.GetString(d, "source_file", "");
        p.Sheet = JsonHelpers.GetString(d, "sheet", "");
        if (d.TryGetProperty("axis", out var ax))
            p.Axis = AxisMapping.FromDictionary(ax);
        if (d.TryGetProperty("pattern_map", out var pm) && pm.ValueKind == JsonValueKind.Object)
            foreach (var kv in pm.EnumerateObject())
            {
                var rp = new RowPattern();
                if (kv.Value.ValueKind == JsonValueKind.String)
                {
                    // Profile cu (1 cot): chuoi pattern -> cot trai.
                    rp.Left = kv.Value.GetString() ?? "";
                }
                else if (kv.Value.ValueKind == JsonValueKind.Object)
                {
                    rp.Left = JsonHelpers.GetString(kv.Value, "left", "");
                    rp.Right = JsonHelpers.GetString(kv.Value, "right", "");
                }
                p.PatternMap[kv.Name] = rp;
            }
        // Cac khoa cu (point_alias, create_missing_pattern, reversible_mode,
        // range_policy) duoc bo qua — tinh nang tuong ung da bo (user chot 08/11).
        p.ReplaceExisting = JsonHelpers.GetBool(d, "replace_existing", true);
        p.SkipZeroRows = JsonHelpers.GetBool(d, "skip_zero_rows", true);
        return p;
    }

    public void Save(string path)
    {
        var options = new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping, // giu nguyen tieng Viet / ° / ² nhu Python ensure_ascii=False
        };
        File.WriteAllText(path, JsonSerializer.Serialize(ToDictionary(), options));
    }

    public static Profile Load(string path)
    {
        using var doc = JsonDocument.Parse(File.ReadAllText(path));
        return FromDictionary(doc.RootElement);
    }
}

internal static class JsonHelpers
{
    public static string GetString(JsonElement d, string key, string fallback) =>
        d.TryGetProperty(key, out var v) && v.ValueKind == JsonValueKind.String
            ? v.GetString() ?? fallback
            : fallback;

    public static int GetInt(JsonElement d, string key, int fallback)
    {
        if (d.TryGetProperty(key, out var v))
        {
            if (v.ValueKind == JsonValueKind.Number && v.TryGetInt32(out int i))
                return i;
            if (v.ValueKind == JsonValueKind.Number)
                return (int)v.GetDouble();
        }
        return fallback;
    }

    public static bool GetBool(JsonElement d, string key, bool fallback) =>
        d.TryGetProperty(key, out var v) && v.ValueKind is JsonValueKind.True or JsonValueKind.False
            ? v.GetBoolean()
            : fallback;
}
