#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace EtabsLoadAssigner.Core;

/// <summary>Loi tu ETABS (ret != 0 hoac joint/pattern khong ton tai).</summary>
public sealed class EtabsError : Exception
{
    public EtabsError(string message) : base(message) { }
}

/// <summary>
/// Giao dien toi thieu ma ca MockEtabs (test offline) va RealEtabs (plugin,
/// goi thang CSI API tren SapModel song) deu thuc hien. Plugin nhan SapModel
/// truc tiep tu cPlugin.Main — KHONG can attach COM nhu app Python chay ngoai.
/// </summary>
public interface IEtabsApi
{
    List<PatternDef> ListPatterns();
    List<string> PatternNames();
    List<PointInfo> ListPoints();
    List<PointInfo> GetSelectedPoints();
    bool CreatePattern(string name, string typeStr = "Other", int? typeCode = null);
    bool AssignForce(string joint, string pattern, double[] values, bool replace = true);
    /// <summary>Xoa tai cua 1 pattern tai joint. Tra false neu khong co tai de xoa (hoac loi).</summary>
    bool DeleteLoadForce(string joint, string pattern);
    /// <summary>
    /// Xoa trang MOI tai tai joint (tat ca pattern trong model) — de diem chi
    /// con tai do plugin gan vao sau do (user chot 08/11). Tra so pattern da
    /// xoa duoc (pattern khong co tai thi bo qua).
    /// </summary>
    int ClearJointLoads(string joint);
    bool RefreshView();
}

public sealed class PointInfo
{
    public string Name { get; set; } = ""; // UniqueName trong ETABS
    public string Label { get; set; } = "";
    public double X { get; set; }
    public double Y { get; set; }
    public double Z { get; set; }
}

/// <summary>
/// Ma loai pattern theo CSI eLoadPatternType — VERIFIED truc tiep tu file
/// CSiAPIv1.dll that (ETABS v22 cua user, enum eLoadPatternType, 2026-11-08);
/// ETABSv1.dll dung chung API nen ma so giong het — se xac minh lai khi co
/// file ETABSv1.dll that.
/// Ban Python cu ghi sai 3 ma (da sua ngay 08/11): Roof Live=4 (thuc te
/// ReduceLive=4, Rooflive=11), Temperature=8 (thuc te 10), Notional=9
/// (thuc te 12 — 9 la Move). Chi co Other=8 la dung tu dau.
/// </summary>
public static class LoadTypeCode
{
    public static readonly IReadOnlyDictionary<string, int> Map = new Dictionary<string, int>
    {
        ["Dead"] = 1,
        ["Super Dead"] = 2,
        ["Live"] = 3,
        ["Roof Live"] = 11, // eLoadPatternType.Rooflive
        ["Seismic"] = 5,    // eLoadPatternType.Quake
        ["Quake"] = 5,
        ["Wind"] = 6,
        ["Snow"] = 7,
        ["Other"] = 8,
        ["Notional"] = 12,
        ["Temperature"] = 10,
    };

    public static int For(string typeStr) => Map.TryGetValue(typeStr, out int code) ? code : 8;
}

/// <summary>
/// Ma loai de tao pattern moi — port cua create_type_for() / _type_to_code() /
/// _CODE_TO_NAME trong etabs_api.py (Python).
///
///   - Pattern _NEG (chieu am cua hang +/-) ke thua dung loai cua pattern goc
///     trong model (EX_NEG &lt;- EX 'Seismic', NX_NEG &lt;- NX 'Notional',
///     T_NEG &lt;- T 'Temperature'=10); pattern khac hoac khong tim thay
///     goc -&gt; Other=8. Quy tac ke thua nay giong het ban Python sau khi
///     sua ma enum ngay 08/11 (truoc do ban Python ghi sai 3 ma).
///   - Ma so dung enum eLoadPatternType (da verify tren CSiAPIv1.dll that;
///     ETABSv1.dll dung chung API — se xac minh lai khi co file):
///     Notional=12 (9 la Move!), Rooflive=11 (4 la ReduceLive),
///     Temperature=10. Chi co Other=8 la giong ban Python cu.
/// </summary>
public static class PatternTypes
{
    public const string NegSuffix = "_NEG";

    /// <summary>Chuan hoa loai pattern (ten 'Seismic' hay ma so '5') ve ma int. Khong biet -> Other=8.</summary>
    public static int TypeToCode(string? typeStr)
    {
        if (typeStr is null)
            return 8;
        string s = typeStr.Trim();
        if (s.Length > 0 && s.All(char.IsDigit))
            return int.Parse(s, System.Globalization.CultureInfo.InvariantCulture);
        return LoadTypeCode.For(s);
    }

    /// <summary>Ma int -> ten loai (nguoc cua LoadTypeCode.Map). Khong biet -> "Other".</summary>
    public static string CodeToName(int code) => code switch
    {
        1 => "Dead",
        2 => "Super Dead",
        3 => "Live",
        5 => "Seismic",
        6 => "Wind",
        7 => "Snow",
        8 => "Other",
        10 => "Temperature",
        11 => "Roof Live",
        12 => "Notional",
        _ => "Other",
    };

    /// <summary>Ma loai de tao 1 pattern moi (xem ghi chu o class).</summary>
    public static int CreateTypeFor(string patternName, IEnumerable<PatternDef> existingPatterns)
    {
        if (patternName.EndsWith(NegSuffix, StringComparison.Ordinal))
        {
            string baseName = patternName.Substring(0, patternName.Length - NegSuffix.Length);
            foreach (var p in existingPatterns)
            {
                if (p.Name == baseName)
                    return TypeToCode(p.Type);
            }
        }
        return 8;
    }
}
/// Seed pattern + diem lay tu anh user gui (29 pattern, diem 20/D...).
/// Port cua MockEtabs trong etabs_api.py.
/// </summary>
public sealed class MockEtabs : IEtabsApi
{
    public static readonly IReadOnlyList<(string Name, string Type, double SelfWeight)> SeedPatterns =
        new List<(string, string, double)>
        {
            ("BC1", "Live", 0), ("BC2", "Live", 0), ("BT1", "Live", 0),
            ("BT2", "Live", 0), ("DLS", "Super Dead", 0), ("EX", "Seismic", 0),
            ("EY", "Seismic", 0), ("EZ", "Seismic", 0), ("F", "Live", 0),
            ("LL1", "Live", 0), ("LL2", "Live", 0), ("LL3", "Live", 0),
            ("LL4", "Live", 0), ("LL5", "Live", 0), ("LLR", "Roof Live", 0),
            ("NX", "Notional", 0), ("NY", "Notional", 0), ("SLL1", "Other", 0),
            ("SLL2", "Other", 0), ("SPL1", "Other", 0), ("SPL2", "Other", 0),
            ("SW", "Dead", 1), ("T", "Temperature", 0), ("WX-", "Wind", 0),
            ("WX+", "Wind", 0), ("WY-1", "Wind", 0), ("WY-2", "Wind", 0),
            ("WY+1", "Wind", 0), ("WY+2", "Wind", 0),
        };

    public List<PatternDef> Patterns { get; } =
        SeedPatterns.Select(p => new PatternDef { Name = p.Name, IsAuto = "No", Type = p.Type, SelfWeight = p.SelfWeight }).ToList();

    public Dictionary<string, PointInfo> Points { get; } = new();

    /// <summary>list[(joint, pattern, [6 gia tri])]</summary>
    public List<(string Joint, string Pattern, double[] Values)> Assignments { get; } = new();

    public List<string> Selected { get; set; } = new() { "20 / D" };

    public bool Connected { get; } = true;

    public MockEtabs()
    {
        foreach (string g in new[] { "20", "27" })
            foreach (string a in new[] { "D", "E" })
            {
                string name = $"{g} / {a}";
                Points[name] = new PointInfo { Name = name, Label = name };
            }
    }

    public List<PatternDef> ListPatterns() => Patterns.ToList();

    public List<string> PatternNames() => Patterns.Select(p => p.Name).ToList();

    public List<PointInfo> ListPoints() => Points.Values.ToList();

    public List<PointInfo> GetSelectedPoints() =>
        Selected.Where(s => Points.ContainsKey(s)).Select(s => Points[s]).ToList();

    public bool CreatePattern(string name, string typeStr = "Other", int? typeCode = null)
    {
        if (PatternNames().Contains(name))
            return false;
        if (typeCode is not null)
            typeStr = PatternTypes.CodeToName(typeCode.Value);
        Patterns.Add(new PatternDef { Name = name, IsAuto = "No", Type = typeStr, SelfWeight = 0 });
        return true;
    }

    public bool AssignForce(string joint, string pattern, double[] values, bool replace = true)
    {
        if (!Points.ContainsKey(joint))
            throw new EtabsError($"Mock: không có joint '{joint}' (có: [{string.Join(", ", Points.Keys.OrderBy(k => k))}]).");
        if (!PatternNames().Contains(pattern))
            throw new EtabsError($"Mock: không có pattern '{pattern}'. Tick 'tự tạo pattern' hoặc match lại.");
        if (replace)
            Assignments.RemoveAll(a => a.Joint == joint && a.Pattern == pattern);
        Assignments.Add((joint, pattern, values.ToArray()));
        return true;
    }

    public bool DeleteLoadForce(string joint, string pattern)
    {
        int removed = Assignments.RemoveAll(a => a.Joint == joint && a.Pattern == pattern);
        return removed > 0;
    }

    public int ClearJointLoads(string joint)
    {
        int removed = Assignments.RemoveAll(a => a.Joint == joint);
        return removed;
    }

    public double[]? GetAssignment(string joint, string pattern)
    {
        foreach (var (j, p, v) in Assignments)
            if (j == joint && p == pattern)
                return v;
        return null;
    }

    public bool RefreshView() => true;
}
