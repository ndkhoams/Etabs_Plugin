#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace EtabsLoadAssigner.Core;

/// <summary>
/// Build danh sach gan tai tu bang da nhan dien + mapping.
/// Port truc tiep tu etabs_loader/assigner.py.
/// Dau ra: list[Assignment] san sang dua vao ETABS (PointObj.SetLoadForce)
/// hoac vao MockEtabs de test offline. Khong tu y doi dau / doi pattern:
/// moi gia tri la, khoang, hang +/- deu ghi chu ro trong Assignment.Notes.
/// </summary>
public static class Assigner
{
    /// <summary>Gan tai cho 1 diem Excel (vd '20 / D') sang 1 joint ETABS.</summary>
    public static BuildResult BuildAssignments(FoundationTable table, string excelPoint, Profile profile)
    {
        var res = new BuildResult();
        if (!table.Points.Contains(excelPoint))
        {
            res.Skipped.Add(("diem",
                $"Diem '{excelPoint}' khong co trong bang (co: [{string.Join(", ", table.Points.Select(p => $"'{p}'"))}])."));
            return res;
        }
        string etabsJoint = profile.PointAlias.TryGetValue(excelPoint, out string? alias) ? alias : excelPoint;
        var axis = profile.Axis;
        var errs = axis.Validate();
        if (errs.Count > 0)
        {
            res.Warnings.AddRange(errs);
            return res;
        }

        foreach (var row in table.Rows)
        {
            string pattern = (profile.PatternFor(row.Name) ?? "").Trim();
            if (pattern == "")
            {
                res.Skipped.Add(("chua match pattern", $"'{row.Name}' o {excelPoint}: chua gan pattern ETABS."));
                continue;
            }
            row.Values.TryGetValue(excelPoint, out var comp);
            comp ??= new Dictionary<string, LoadValue>();
            comp.TryGetValue("Hx", out var hxLv);
            comp.TryGetValue("Hy", out var hyLv);
            comp.TryGetValue("Vz", out var vzLv);
            if (hxLv is null || hyLv is null || vzLv is null)
            {
                res.Skipped.Add(("thieu so", $"'{row.Name}' o {excelPoint}: thieu o Hx/Hy/Vz."));
                continue;
            }
            var notes = new List<string>();
            if (row.Reversible)
                notes.Add("hang +/-, co the dao chieu");
            var vals = new List<double>();
            bool ok = true;
            foreach (var lv in new[] { hxLv, hyLv, vzLv })
            {
                if (lv.Kind == LoadValueKind.Text)
                {
                    res.Skipped.Add(("o khong phai so", $"'{row.Name}' o {excelPoint}: o '{lv.Raw}' khong phai so."));
                    ok = false;
                    break;
                }
                if (lv.Kind == LoadValueKind.Blank)
                {
                    vals.Add(0.0);
                }
                else if (lv.Kind == LoadValueKind.Range)
                {
                    var (v, note) = Mapping.ResolveRange(lv, profile.RangePolicy);
                    if (v is null)
                    {
                        res.Skipped.Add(("khoang", $"'{row.Name}' o {excelPoint}: {note}."));
                        ok = false;
                        break;
                    }
                    vals.Add(v.Value);
                    notes.Add(note);
                }
                else
                {
                    vals.Add(lv.Value!.Value);
                }
            }
            if (!ok)
                continue;
            var (fx, fy, fz) = axis.Apply(vals[0], vals[1], vals[2]);
            var baseAssignment = new Assignment
            {
                ExcelPoint = excelPoint,
                EtabsJoint = etabsJoint,
                ExcelRow = row.Name,
                Pattern = pattern,
                Fx = fx, Fy = fy, Fz = fz,
                Notes = string.Join("; ", notes),
            };
            if (profile.SkipZeroRows && baseAssignment.IsZero())
            {
                res.Skipped.Add(("dong 0", $"'{row.Name}' o {excelPoint}: ca 3 luc = 0, bo qua."));
                continue;
            }
            res.Assignments.Add(baseAssignment);
            if (row.Reversible && profile.ReversibleMode == "both")
            {
                var neg = new Assignment
                {
                    ExcelPoint = excelPoint,
                    EtabsJoint = etabsJoint,
                    ExcelRow = row.Name,
                    Pattern = pattern + "_NEG",
                    Fx = -fx, Fy = -fy, Fz = -fz,
                    Negated = true,
                    Notes = "ban dao chieu cua hang +/- (pattern _NEG)",
                };
                if (!(profile.SkipZeroRows && neg.IsZero()))
                {
                    res.Assignments.Add(neg);
                    res.Warnings.Add(
                        $"'{row.Name}': tao them pattern '{pattern}_NEG' cho chieu am — " +
                        "nho kiem tra pattern nay co trong model khong.");
                }
            }
        }
        return res;
    }

    public static string PreviewText(FoundationTable table, string excelPoint, Profile profile, BuildResult result)
    {
        string joint = profile.PointAlias.TryGetValue(excelPoint, out string? alias) ? alias : excelPoint;
        var lines = new List<string>
        {
            $"Diem Excel: {excelPoint}  ->  Joint ETABS: {joint}",
            $"Phuong/chieu: {profile.Axis.Describe()}   (don vi kN, M = 0 — bang goc chi co luc)",
            $"So luong gan: {result.Assignments.Count} | Bo qua: {result.Skipped.Count} | Canh bao: {result.Warnings.Count}",
            "",
            $"{"Pattern",-10} {"Fx",8} {"Fy",8} {"Fz",8}   Nguon (dong Excel)",
        };
        foreach (var a in result.Assignments)
        {
            string extra = a.Notes != "" ? $"  [{a.Notes}]" : "";
            lines.Add($"{a.Pattern,-10} {a.Fx,8:F2} {a.Fy,8:F2} {a.Fz,8:F2}   {a.ExcelRow}{extra}");
        }
        lines.Add("");
        lines.Add("Tong theo pattern:");
        foreach (var kv in result.TotalsByPattern().OrderBy(k => k.Key))
            lines.Add($"  {kv.Key,-10} Fx={kv.Value.Fx,8:F2}  Fy={kv.Value.Fy,8:F2}  Fz={kv.Value.Fz,8:F2}  ({kv.Value.Count} dong)");
        if (result.Skipped.Count > 0)
        {
            lines.Add("");
            lines.Add("Da bo qua:");
            foreach (var (kind, msg) in result.Skipped)
                lines.Add($"  - [{kind}] {msg}");
        }
        if (result.Warnings.Count > 0)
        {
            lines.Add("");
            lines.Add("Canh bao:");
            foreach (var w in result.Warnings)
                lines.Add($"  ! {w}");
        }
        return string.Join("\n", lines);
    }
}

public sealed class Assignment
{
    public string ExcelPoint { get; set; } = "";
    public string EtabsJoint { get; set; } = "";
    public string ExcelRow { get; set; } = "";
    public string Pattern { get; set; } = "";
    public double Fx { get; set; }
    public double Fy { get; set; }
    public double Fz { get; set; }
    public double Mx { get; set; }
    public double My { get; set; }
    public double Mz { get; set; }
    public string Notes { get; set; } = "";
    public bool Negated { get; set; } // Ban dao chieu cua hang +/-.

    public double[] Values() => new[] { Fx, Fy, Fz, Mx, My, Mz };

    public bool IsZero() => Values().All(v => v == 0);
}

public sealed class BuildResult
{
    public List<Assignment> Assignments { get; set; } = new();
    public List<(string Reason, string Message)> Skipped { get; set; } = new();
    public List<string> Warnings { get; set; } = new();

    public Dictionary<string, (double Fx, double Fy, double Fz, int Count)> TotalsByPattern()
    {
        var tot = new Dictionary<string, (double Fx, double Fy, double Fz, int Count)>();
        foreach (var a in Assignments)
        {
            if (!tot.TryGetValue(a.Pattern, out var t))
                t = (0, 0, 0, 0);
            tot[a.Pattern] = (t.Fx + a.Fx, t.Fy + a.Fy, t.Fz + a.Fz, t.Count + 1);
        }
        return tot;
    }
}
