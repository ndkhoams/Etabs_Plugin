using ETABSv1;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Etabs_Ultimate_Tools
{
    /// <summary>
    /// Các hàm tiện ích dùng chung, tránh trùng lặp code giữa các class.
    /// </summary>
    internal static class EtabsHelper
    {
        public class StoryInfo
        {
            public string Name { get; set; }
            public double Elevation { get; set; }
            public double Height { get; set; }
        }

        /// <summary>Đọc danh sách tầng từ ETABS, sắp xếp theo cao độ tăng dần.</summary>
        public static List<StoryInfo> ReadStories(cSapModel sap)
        {
            int n = 0;
            string[] names = null;
            double[] elevations = null, heights = null;
            bool[] isMaster = null, spliceAbove = null;
            string[] similarTo = null;
            double[] spliceHeight = null;

            int ret = sap.Story.GetStories(ref n, ref names, ref elevations, ref heights,
                ref isMaster, ref similarTo, ref spliceAbove, ref spliceHeight);
            if (ret != 0 || n < 0 || names == null || elevations == null ||
                names.Length < n || elevations.Length < n)
                throw new InvalidOperationException(
                    "ETABS Story.GetStories thất bại hoặc trả về mảng tầng không hợp lệ (return code " + ret + ").");

            var list = new List<StoryInfo>(n);
            for (int i = 0; i < n; i++)
                list.Add(new StoryInfo
                {
                    Name = names[i],
                    Elevation = elevations[i],
                    Height = heights != null && i < heights.Length ? heights[i] : 0.0
                });

            return list.OrderBy(x => x.Elevation).ToList();
        }

        /// <summary>Bổ sung tầng có trong kết quả Story Drifts nhưng thiếu trong GetStories.</summary>
        public static void AddMissingStoriesFromDrifts(
            cSapModel sap, List<StoryInfo> stories, params string[] loadCases)
        {
            var knownStories = new HashSet<string>(
                stories.Select(s => s.Name), StringComparer.OrdinalIgnoreCase);

            foreach (string loadCase in loadCases
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                try
                {
                    SelectCaseOrCombo(sap, loadCase);

                    int count = 0;
                    string[] storyNames = null, caseNames = null, stepTypes = null;
                    string[] directions = null, labels = null;
                    double[] stepNumbers = null, drifts = null, x = null, y = null, z = null;
                    int ret = sap.Results.StoryDrifts(ref count, ref storyNames, ref caseNames,
                        ref stepTypes, ref stepNumbers, ref directions, ref drifts,
                        ref labels, ref x, ref y, ref z);
                    if (ret != 0 || storyNames == null) continue;

                    for (int i = 0; i < count && i < storyNames.Length; i++)
                    {
                        string storyName = storyNames[i];
                        if (string.IsNullOrWhiteSpace(storyName) || knownStories.Contains(storyName))
                            continue;
                        if (caseNames != null && i < caseNames.Length &&
                            !string.Equals(caseNames[i], loadCase, StringComparison.OrdinalIgnoreCase))
                            continue;

                        double elevation = 0.0, height = 0.0;
                        if (sap.Story.GetElevation(storyName, ref elevation) != 0) continue;
                        sap.Story.GetHeight(storyName, ref height);

                        stories.Add(new StoryInfo
                        {
                            Name = storyName,
                            Elevation = elevation,
                            Height = height
                        });
                        knownStories.Add(storyName);
                    }
                }
                catch
                {
                    // Keep the original story list if this result set is unavailable.
                }
            }

            stories.Sort((a, b) => a.Elevation.CompareTo(b.Elevation));
        }

        /// <summary>Trả về true nếu tên tầng là "Base" hoặc chứa "Base".</summary>
        public static bool IsBaseLevel(string storyName)
        {
            if (string.IsNullOrWhiteSpace(storyName)) return false;
            return storyName.Trim().IndexOf("Base", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        /// <summary>So sánh tên output case sau khi trim, không phân biệt hoa thường.</summary>
        public static bool IsSameOrBlank(string outputCase, string selectedName)
        {
            return !string.IsNullOrWhiteSpace(outputCase)
                && !string.IsNullOrWhiteSpace(selectedName)
                && string.Equals(outputCase.Trim(), selectedName.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        public static bool IsSameOrEnvelopeCase(string outputCase, string selectedName)
        {
            if (IsSameOrBlank(outputCase, selectedName)) return true;
            if (string.IsNullOrWhiteSpace(outputCase) || string.IsNullOrWhiteSpace(selectedName))
                return false;

            string actual = outputCase.Trim();
            string selected = selectedName.Trim();
            string[] suffixes = { " (Max)", " (Min)", "_max", "_min", "-max", "-min", " Max", " Min" };
            foreach (string suffix in suffixes)
            {
                if (actual.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(actual.Substring(0, actual.Length - suffix.Length).TrimEnd(),
                        selected, StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            return false;
        }

        public static void SelectCaseOrCombo(cSapModel sap, string name)
        {
            sap.Results.Setup.DeselectAllCasesAndCombosForOutput();
            int ret = sap.Results.Setup.SetCaseSelectedForOutput(name);
            if (ret != 0) sap.Results.Setup.SetComboSelectedForOutput(name);
        }

        public static void SelectComboOnly(cSapModel sap, string comboName)
        {
            sap.Results.Setup.DeselectAllCasesAndCombosForOutput();
            sap.Results.Setup.SetComboSelectedForOutput(comboName);
        }

        public static List<string> GetLoadCombinations(cSapModel sap)
        {
            int n = 0;
            string[] names = null;
            sap.RespCombo.GetNameList(ref n, ref names);
            return names == null ? new List<string>() : names.OrderBy(x => x).ToList();
        }

        /// <summary>Áp dụng thiết lập trang in chuẩn A4 cho một worksheet.</summary>
        public static void ApplyA4PageSetup(ClosedXML.Excel.IXLWorksheet ws)
        {
            ws.PageSetup.PaperSize = ClosedXML.Excel.XLPaperSize.A4Paper;
            ws.PageSetup.PageOrientation = ClosedXML.Excel.XLPageOrientation.Portrait;
            ws.PageSetup.FitToPages(1, 0);
            ws.PageSetup.Margins.Left = 0.75;
            ws.PageSetup.Margins.Right = 0.75;
            ws.PageSetup.Margins.Top = 0.75;
            ws.PageSetup.Margins.Bottom = 0.50;
            ws.PageSetup.Margins.Header = 0.50;
            ws.PageSetup.Margins.Footer = 0.75;
            ws.PageSetup.CenterHorizontally = true;

            ws.Style.Font.FontName = "Arial";
            ws.Style.Font.FontSize = 11;
        }

        /// <summary>
        /// Parse số thập phân KHÔNG phụ thuộc culture của Windows (double.TryParse mặc định
        /// theo CurrentCulture: trên Windows tiếng Việt "3.12" bị hiểu thành 312).
        /// Chấp nhận "3.12", "3,12" và kiểu VN "1.234,5" (chấm = phân cách nghìn, phẩy = thập phân).
        /// Chuỗi chỉ gồm các nhóm nghìn kiểu "1.000" được hiểu là 1000.
        /// </summary>
        /// <summary>
        /// Đọc Story Drifts cho một case/combo + phương, trả về map tầng -> max |drift|.
        /// Hàm dùng chung cho cả 3 extractor P-Delta / Seismic / Wind Drift (trước đây mỗi
        /// file giữ một bản sao riêng). API lỗi hoặc mảng thiếu -> map rỗng (nơi gọi sẽ
        /// đánh dấu tầng là "Không có dữ liệu").
        /// </summary>
        public static Dictionary<string, double> ReadStoryDrifts(cSapModel sap, string loadCase, string dir)
        {
            SelectCaseOrCombo(sap, loadCase);

            int n = 0;
            string[] story = null, caseName = null, stepType = null, dirArr = null, label = null;
            double[] stepNum = null, drift = null, x = null, y = null, z = null;
            int ret = sap.Results.StoryDrifts(ref n, ref story, ref caseName, ref stepType,
                ref stepNum, ref dirArr, ref drift, ref label, ref x, ref y, ref z);

            var result = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            if (ret != 0 || story == null || caseName == null || dirArr == null || drift == null)
                return result;

            int count = Math.Min(n, Math.Min(story.Length,
                Math.Min(caseName.Length, Math.Min(dirArr.Length, drift.Length))));
            for (int i = 0; i < count; i++)
            {
                if (story[i] == null || caseName[i] == null || dirArr[i] == null) continue;
                if (!string.Equals(caseName[i], loadCase, StringComparison.OrdinalIgnoreCase)) continue;
                if (!dirArr[i].StartsWith(dir, StringComparison.OrdinalIgnoreCase)) continue;

                double val = Math.Abs(drift[i]);
                if (!result.TryGetValue(story[i], out var cur) || val > cur)
                    result[story[i]] = val;
            }
            return result;
        }

        public static bool TryParseNumber(string text, out double value)
        {
            value = 0.0;
            if (string.IsNullOrWhiteSpace(text)) return false;
            string s = text.Trim();
            try
            {
                bool hasDot = s.IndexOf('.') >= 0;
                bool hasComma = s.IndexOf(',') >= 0;
                if (hasDot && hasComma)
                    s = s.Replace(".", "").Replace(",", ".");
                else if (hasComma)
                    s = s.Replace(",", ".");
                else if (hasDot && System.Text.RegularExpressions.Regex.IsMatch(
                    s, @"^\d{1,3}(\.\d{3})+$"))
                    s = s.Replace(".", "");

                return double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
            }
            catch
            {
                value = 0.0;
                return false;
            }
        }
    }

    /// <summary>
    /// Chụp trạng thái toàn cục của phiên ETABS mà các tính năng của tool tạm thay đổi —
    /// đơn vị hiện hành, lựa chọn case/combo xuất kết quả (Results.Setup), tùy chọn combo
    /// nhiều giá trị, lựa chọn hiển thị của Database Tables — và khôi phục khi Dispose.
    /// Dùng: using (new EtabsStateGuard(sap)) { ... } quanh mỗi lần chạy tính năng.
    /// Mọi bước đều best-effort (try/catch riêng): lỗi một bước không chặn các bước còn lại,
    /// và việc chụp trạng thái không bao giờ làm hỏng lần chạy chính.
    /// </summary>
    public sealed class EtabsStateGuard : IDisposable
    {
        private readonly cSapModel _sap;
        private eUnits _units;
        private bool _hasUnits;
        private readonly List<string> _cases = new List<string>();
        private readonly List<string> _combos = new List<string>();
        private bool _hasSelection;
        private int _multiValuedOption;
        private bool _hasMultiValuedOption;
        private string[] _dbCases;
        private bool _hasDbCases;
        private string[] _dbCombos;
        private bool _hasDbCombos;
        private bool _disposed;

        public EtabsStateGuard(cSapModel sap)
        {
            _sap = sap;
            if (_sap == null) return;

            try
            {
                // ETABS 22: GetPresentUnits() không có tham số, trả eUnits trực tiếp.
                _units = _sap.GetPresentUnits();
                _hasUnits = true;
            }
            catch { }

            try
            {
                int n = 0;
                string[] names = null;
                if (_sap.LoadCases.GetNameList(ref n, ref names) == 0 && names != null)
                {
                    foreach (string name in names)
                    {
                        bool sel = false;
                        try
                        {
                            if (_sap.Results.Setup.GetCaseSelectedForOutput(name, ref sel) == 0 && sel)
                                _cases.Add(name);
                        }
                        catch { }
                    }
                }
                n = 0; names = null;
                if (_sap.RespCombo.GetNameList(ref n, ref names) == 0 && names != null)
                {
                    foreach (string name in names)
                    {
                        bool sel = false;
                        try
                        {
                            if (_sap.Results.Setup.GetComboSelectedForOutput(name, ref sel) == 0 && sel)
                                _combos.Add(name);
                        }
                        catch { }
                    }
                }
                _hasSelection = true;
            }
            catch { }

            try
            {
                int opt = 0;
                if (_sap.Results.Setup.GetOptionMultiValuedCombo(ref opt) == 0)
                { _multiValuedOption = opt; _hasMultiValuedOption = true; }
            }
            catch { }

            try
            {
                int n = 0;
                string[] list = null;
                if (_sap.DatabaseTables.GetLoadCasesSelectedForDisplay(ref n, ref list) == 0)
                { _dbCases = list; _hasDbCases = true; }
            }
            catch { }

            try
            {
                int n = 0;
                string[] list = null;
                if (_sap.DatabaseTables.GetLoadCombinationsSelectedForDisplay(ref n, ref list) == 0)
                { _dbCombos = list; _hasDbCombos = true; }
            }
            catch { }
        }

        public void Dispose()
        {
            if (_disposed || _sap == null) return;
            _disposed = true;

            if (_hasSelection)
            {
                try { _sap.Results.Setup.DeselectAllCasesAndCombosForOutput(); } catch { }
                foreach (string name in _cases)
                {
                    try { _sap.Results.Setup.SetCaseSelectedForOutput(name, true); } catch { }
                }
                foreach (string name in _combos)
                {
                    try { _sap.Results.Setup.SetComboSelectedForOutput(name, true); } catch { }
                }
            }

            if (_hasMultiValuedOption)
            {
                try { _sap.Results.Setup.SetOptionMultiValuedCombo(_multiValuedOption); } catch { }
            }

            if (_hasDbCases)
            {
                try
                {
                    string[] list = _dbCases ?? new string[0];
                    _sap.DatabaseTables.SetLoadCasesSelectedForDisplay(ref list);
                }
                catch { }
            }

            if (_hasDbCombos)
            {
                try
                {
                    string[] list = _dbCombos ?? new string[0];
                    _sap.DatabaseTables.SetLoadCombinationsSelectedForDisplay(ref list);
                }
                catch { }
            }

            if (_hasUnits)
            {
                try { _sap.SetPresentUnits(_units); } catch { }
            }
        }
    }
}
