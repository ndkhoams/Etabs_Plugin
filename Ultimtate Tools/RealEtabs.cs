#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ETABSv1;
using EtabsLoadAssigner.Core;

namespace EtabsLoadAssigner.Plugin;

/// <summary>
/// Ket noi ETABS that qua ETABSv1 tren SapModel song do cPlugin.Main nhan.
/// KHONG attach COM, KHONG mo ETABS moi — model dang mo san.
///
/// MOI CHU KY HAM duoi day lay truc tiep tu file CSiAPIv1.dll that
/// (ETABS v22 cua user) truoc day, va API ETABSv1 co cung chu ky —
/// can xac minh lai enum eLoadPatternType tu ETABSv1.dll that khi co file:
///   - SapModel.SelectObj (khong phai ".Select" nhu comtypes goi tat)
///   - LoadPatterns.GetLoadType chi tra ve MyType (khong co 4 gia tri
///     nhu code Python unpack — SelfWTMultiplier lay rieng bang
///     GetSelfWTMultiplier)
///   - PointObj.GetLabelFromName(Name, ref Label, ref Story) — chi 2 out
///   - PointObj.SetLoadForce(..., Replace = False MAC DINH) — phai truyen ro
///   - PointObj.GetAllPoints / GetLabelNameList: lay het 1 lenh
///   - eLoadPatternType (da verify tren CSiAPIv1.dll): Rooflive=11,
///     Temperature=10, Notional=12
///     (ban Python ghi sai 4/11/8/9 — xem LoadTypeCode)
/// </summary>
public sealed class RealEtabs : IEtabsApi
{
    private readonly cSapModel _sap;

    public RealEtabs(cSapModel sapModel)
    {
        _sap = sapModel;
        // Don vi truyen qua API: kN, m, C (eUnits.kN_m_C = 6, giong app Python).
        try { _sap.SetPresentUnits(eUnits.kN_m_C); } catch { /* bo qua — model giu don vi hien tai */ }
    }

    private static void CheckRet(int ret, string what)
    {
        if (ret != 0)
            throw new EtabsError($"ETABS tra ret={ret} khi {what} (0 = OK). " +
                "Doi chieu 'CSI API ETABS v1.chm' trong thu muc cai ETABS.");
    }

    public List<PatternDef> ListPatterns()
    {
        int n = 0;
        string[] names = Array.Empty<string>();
        CheckRet(_sap.LoadPatterns.GetNameList(ref n, ref names), "lay danh sach pattern");
        var list = new List<PatternDef>();
        foreach (string nm in names ?? Array.Empty<string>())
        {
            var type = eLoadPatternType.Other;
            string typeStr = "Other";
            try
            {
                if (_sap.LoadPatterns.GetLoadType(nm, ref type) == 0)
                    typeStr = type.ToString();
            }
            catch { /* giu "Other" */ }
            double sw = 0;
            try { _sap.LoadPatterns.GetSelfWTMultiplier(nm, ref sw); } catch { /* giu 0 */ }
            list.Add(new PatternDef { Name = nm, Type = typeStr, SelfWeight = sw });
        }
        return list;
    }

    public List<string> PatternNames() => ListPatterns().Select(p => p.Name).ToList();

    public List<PointInfo> ListPoints()
    {
        // Mot lenh lay het ten + toa do.
        int n = 0;
        string[] names = Array.Empty<string>();
        double[] xs = Array.Empty<double>(), ys = Array.Empty<double>(), zs = Array.Empty<double>();
        if (_sap.PointObj.GetAllPoints(ref n, ref names, ref xs, ref ys, ref zs) == 0 && names is not null)
        {
            var labels = GetLabels(names);
            var pts = new List<PointInfo>();
            for (int i = 0; i < names.Length; i++)
                pts.Add(new PointInfo
                {
                    Name = names[i],
                    Label = labels.TryGetValue(names[i], out string? lb) ? lb : names[i],
                    X = i < xs.Length ? xs[i] : 0,
                    Y = i < ys.Length ? ys[i] : 0,
                    Z = i < zs.Length ? zs[i] : 0,
                });
            return pts;
        }
        // Fallback: GetNameList + GetCoordCartesian tung diem.
        int n2 = 0;
        string[] names2 = Array.Empty<string>();
        CheckRet(_sap.PointObj.GetNameList(ref n2, ref names2), "lay danh sach point");
        var labels2 = GetLabels(names2 ?? Array.Empty<string>());
        var pts2 = new List<PointInfo>();
        foreach (string nm in names2 ?? Array.Empty<string>())
        {
            double x = 0, y = 0, z = 0;
            try { _sap.PointObj.GetCoordCartesian(nm, ref x, ref y, ref z); } catch { /* giu 0 */ }
            pts2.Add(new PointInfo
            {
                Name = nm,
                Label = labels2.TryGetValue(nm, out string? lb) ? lb : nm,
                X = x, Y = y, Z = z,
            });
        }
        return pts2;
    }

    private Dictionary<string, string> GetLabels(string[] names)
    {
        var map = new Dictionary<string, string>();
        try
        {
            int n = 0;
            string[] nm = Array.Empty<string>(), lb = Array.Empty<string>(), st = Array.Empty<string>();
            if (_sap.PointObj.GetLabelNameList(ref n, ref nm, ref lb, ref st) == 0 && nm is not null)
                for (int i = 0; i < nm.Length; i++)
                    map[nm[i]] = i < lb.Length ? lb[i] : nm[i];
        }
        catch { /* fallback: goi tung diem */ }
        foreach (string name in names)
        {
            if (map.ContainsKey(name))
                continue;
            try
            {
                string label = "", story = "";
                if (_sap.PointObj.GetLabelFromName(name, ref label, ref story) == 0)
                    map[name] = label;
            }
            catch { /* giu ten */ }
        }
        return map;
    }

    public List<PointInfo> GetSelectedPoints()
    {
        // CSI API: SapModel.SelectObj.GetSelected() -> (n, ObjectType[], ObjectName[])
        // ObjectType: 1 = point object (theo quy uoc CSI API; khong co enum trong CSiAPIv1).
        int n = 0;
        int[] types = Array.Empty<int>();
        string[] names = Array.Empty<string>();
        CheckRet(_sap.SelectObj.GetSelected(ref n, ref types, ref names), "lay diem dang chon");
        names ??= Array.Empty<string>();
        types ??= Array.Empty<int>();
        var pts = ListPoints().ToDictionary(p => p.Name);
        var list = new List<PointInfo>();
        for (int i = 0; i < names.Length; i++)
        {
            int t = i < types.Length ? types[i] : -1;
            if (t == 1 && pts.TryGetValue(names[i], out var p))
                list.Add(p);
        }
        return list;
    }

    public bool CreatePattern(string name, string typeStr = "Other", int? typeCode = null)
    {
        // Port cua RealEtabs.create_pattern() (Python): uu tien type_code,
        // fallback ve Other=8 neu Add that bai.
        int code = typeCode ?? LoadTypeCode.For(typeStr);
        var type = (eLoadPatternType)code;
        int ret = _sap.LoadPatterns.Add(name, type, 0, true);
        if (ret != 0 && code != (int)eLoadPatternType.Other)
            ret = _sap.LoadPatterns.Add(name, eLoadPatternType.Other, 0, true);
        CheckRet(ret, $"tao pattern '{name}' (type {typeStr})");
        return true;
    }

    public bool AssignForce(string joint, string pattern, double[] values, bool replace = true)
    {
        // PointObj.SetLoadForce(Name, LoadPat, Value[6], Replace, CSys=Global, ItemType=Objects)
        // Value = [Fx, Fy, Fz, Mx, My, Mz]; Replace=true = ghi de len tai cu
        // cua cung pattern tai diem nay.
        // LUU Y: tham so Replace cua API mac dinh la FALSE — phai truyen ro.
        double[] v = values.ToArray();
        int ret = _sap.PointObj.SetLoadForce(joint, pattern, ref v, replace);
        CheckRet(ret, $"gan tai {pattern} tai joint '{joint}'");
        return true;
    }

    public bool RefreshView()
    {
        try { _sap.View.RefreshView(); } catch { /* bo qua */ }
        return true;
    }
}
