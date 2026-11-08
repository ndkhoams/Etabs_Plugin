#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace EtabsLoadAssigner.Core;

/// <summary>
/// Build danh sach gán tải tu bang da nhan dien + mapping 2 cot.
/// Port truc tiep tu etabs_loader/assigner.py (da doi theo user chot 08/11:
/// bo co che pattern _NEG tu tao + chinh sach khoang; thay bang 2 cot
/// trai/phai do user chon, tool khong tu tao pattern).
/// Dau ra: list[Assignment] san sang dua vao ETABS (PointObj.SetLoadForce)
/// hoac vao MockEtabs de test offline.
/// </summary>
public static class Assigner
{
    /// <summary>
    /// Gan tai cho 1 diem Excel (vd '20 / D') sang 1 joint ETABS (ten joint
    /// that lay tu lua chon trong ETABS — KHONG dung alias go tay nua).
    /// Match 2 cot (user chot 08/11): cot trai = gia tri nho (gia tri don,
    /// hoac min cua o khoang / -V cua hang +/-); cot phai = gia tri lon
    /// (max cua o khoang / +V cua hang +/-). Dong khong co +/- chi can cot
    /// trai. Tool KHONG tu tao pattern — pattern phai co san trong model.
    /// </summary>
    public static BuildResult BuildAssignments(FoundationTable table, string excelPoint, string etabsJoint, Profile profile)
    {
        var res = new BuildResult();
        if (!table.Points.Contains(excelPoint))
        {
            res.Skipped.Add(("diem",
                $"Điểm '{excelPoint}' không có trong bảng (có: [{string.Join(", ", table.Points.Select(p => $"'{p}'"))}])."));
            return res;
        }
        if (etabsJoint == "")
        {
            res.Skipped.Add(("diem",
                "Chưa chọn joint trong ETABS — bấm 'Chọn điểm trong ETABS' để chọn joint cần gán."));
            return res;
        }
        var axis = profile.Axis;
        var errs = axis.Validate();
        if (errs.Count > 0)
        {
            res.Warnings.AddRange(errs);
            return res;
        }

        foreach (var row in table.Rows)
        {
            var rp = profile.PatternFor(row.Name);
            string leftPat = (rp.Left ?? "").Trim();
            string rightPat = (rp.Right ?? "").Trim();
            if (leftPat == "" && rightPat == "")
            {
                res.Skipped.Add(("chua match pattern", $"'{row.Name}' ở {excelPoint}: chưa gán pattern ETABS (cột trái)."));
                continue;
            }
            if (leftPat == "")
            {
                res.Skipped.Add(("chua match pattern", $"'{row.Name}' ở {excelPoint}: chưa chọn pattern cột trái (bắt buộc)."));
                continue;
            }
            row.Values.TryGetValue(excelPoint, out var comp);
            comp ??= new Dictionary<string, LoadValue>();
            comp.TryGetValue("Hx", out var hxLv);
            comp.TryGetValue("Hy", out var hyLv);
            comp.TryGetValue("Vz", out var vzLv);
            if (hxLv is null || hyLv is null || vzLv is null)
            {
                res.Skipped.Add(("thieu so", $"'{row.Name}' ở {excelPoint}: thiếu ô Hx/Hy/Vz."));
                continue;
            }
            // Tach moi thanh phan thanh (nho, lon).
            //  - So don V, dong khong +/-: (V, V).
            //  - So don V, dong +/-: (-V, +V).
            //  - Khoang (a, b): (a, b). - O trong: (0, 0). - Chu: bo dong.
            bool hasPlusMinus = row.Reversible;
            var leftVals = new List<double>();
            var rightVals = new List<double>();
            bool ok = true;
            foreach (var lv in new[] { hxLv, hyLv, vzLv })
            {
                if (lv.Kind == LoadValueKind.Text)
                {
                    res.Skipped.Add(("o khong phai so", $"'{row.Name}' ở {excelPoint}: ô '{lv.Raw}' không phải số."));
                    ok = false;
                    break;
                }
                if (lv.Kind == LoadValueKind.Blank)
                {
                    leftVals.Add(0.0);
                    rightVals.Add(0.0);
                }
                else if (lv.Kind == LoadValueKind.Range)
                {
                    hasPlusMinus = true;
                    leftVals.Add(lv.VMin!.Value);
                    rightVals.Add(lv.VMax!.Value);
                }
                else
                {
                    double v = lv.Value!.Value;
                    if (row.Reversible)
                    {
                        leftVals.Add(-v);
                        rightVals.Add(v);
                    }
                    else
                    {
                        leftVals.Add(v);
                        rightVals.Add(v);
                    }
                }
            }
            if (!ok)
                continue;
            var (fxL, fyL, fzL) = axis.Apply(leftVals[0], leftVals[1], leftVals[2]);
            var (fxR, fyR, fzR) = axis.Apply(rightVals[0], rightVals[1], rightVals[2]);
            var leftAssignment = new Assignment
            {
                ExcelPoint = excelPoint,
                EtabsJoint = etabsJoint,
                ExcelRow = row.Name,
                Pattern = leftPat,
                Fx = fxL, Fy = fyL, Fz = fzL,
                Side = "trai",
                Notes = hasPlusMinus ? "giá trị nhỏ (cột trái)" : "",
            };
            if (profile.SkipZeroRows && leftAssignment.IsZero())
            {
                res.Skipped.Add(("dong 0", $"'{row.Name}' ở {excelPoint}: cả 3 lực = 0, bỏ qua."));
                continue;
            }
            res.Assignments.Add(leftAssignment);
            if (hasPlusMinus)
            {
                if (rightPat == "")
                {
                    res.Warnings.Add(
                        $"'{row.Name}': có giá trị +/- nhưng chưa chọn pattern cột phải — " +
                        "chỉ gán giá trị nhỏ (cột trái), giá trị lớn chưa được gán.");
                }
                else
                {
                    var rightAssignment = new Assignment
                    {
                        ExcelPoint = excelPoint,
                        EtabsJoint = etabsJoint,
                        ExcelRow = row.Name,
                        Pattern = rightPat,
                        Fx = fxR, Fy = fyR, Fz = fzR,
                        Side = "phai",
                        Notes = "giá trị lớn (cột phải)",
                    };
                    if (!(profile.SkipZeroRows && rightAssignment.IsZero()))
                        res.Assignments.Add(rightAssignment);
                }
            }
        }
        return res;
    }

    public static string PreviewText(FoundationTable table, string excelPoint, string etabsJoint, Profile profile, BuildResult result)
    {
        var lines = new List<string>
        {
            $"Điểm Excel: {excelPoint}  ->  Joint ETABS: {etabsJoint}",
            $"Phương/chiều: {profile.Axis.Describe()}   (đơn vị kN, M = 0 — bảng gốc chỉ có lực)",
            $"Số lượng gán: {result.Assignments.Count} | Bỏ qua: {result.Skipped.Count} | Cảnh báo: {result.Warnings.Count}",
            "",
            $"{"Pattern",-10} {"Fx",8} {"Fy",8} {"Fz",8}   Nguon (dong Excel)",
        };
        foreach (var a in result.Assignments)
        {
            string extra = a.Notes != "" ? $"  [{a.Notes}]" : "";
            lines.Add($"{a.Pattern,-10} {a.Fx,8:F2} {a.Fy,8:F2} {a.Fz,8:F2}   {a.ExcelRow}{extra}");
        }
        lines.Add("");
        lines.Add("Tổng theo pattern:");
        foreach (var kv in result.TotalsByPattern().OrderBy(k => k.Key))
            lines.Add($"  {kv.Key,-10} Fx={kv.Value.Fx,8:F2}  Fy={kv.Value.Fy,8:F2}  Fz={kv.Value.Fz,8:F2}  ({kv.Value.Count} dong)");
        if (result.Skipped.Count > 0)
        {
            lines.Add("");
            lines.Add("Đã bỏ qua:");
            foreach (var (kind, msg) in result.Skipped)
                lines.Add($"  - [{kind}] {msg}");
        }
        if (result.Warnings.Count > 0)
        {
            lines.Add("");
            lines.Add("Cảnh báo:");
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
    public string Side { get; set; } = "trai"; // "trai" = gia tri nho / gia tri don; "phai" = gia tri lon (chi khi co +/-).

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
