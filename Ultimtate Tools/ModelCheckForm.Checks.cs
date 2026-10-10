using ETABSv1;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Etabs_Ultimate_Tools
{
    public partial class ModelCheckForm : Form
    {
        // ---------- Tính toán ----------

        private void RunCheck()
        {
            if (!double.TryParse(txtQ.Text, out var q)) q = 1.0;
            _qFactor = q;

            if (!RequireCombo(cboCombo, "Check Model", "Chưa chọn tổ hợp kiểm tra.", out var combo)) return;

            _sap.SetPresentUnits(eUnits.kN_m_C);
            _sap.Results.Setup.DeselectAllCasesAndCombosForOutput();

            _rows = new List<PDeltaCheckRow>();
            var displacementSource = GetDisplacementSource(cboPDeltaSource);
            var aggregationMode = GetDiaphragmAggregationMode(cboPDeltaDiaphragmMode);
            _rows.AddRange(PDeltaExtractor.Calculate(_sap, combo, combo, "X", q, displacementSource, aggregationMode));
            _rows.AddRange(PDeltaExtractor.Calculate(_sap, combo, combo, "Y", q, displacementSource, aggregationMode));
            _rows = _rows.OrderBy(r => r.Direction).ThenByDescending(r => r.Elevation).ToList();

            if (_rows.Count == 0)
                Warn("Không có kết quả P-Delta cho tổ hợp " + combo +
                    ". Kiểm tra bảng Story Forces, Output Case và kết quả phân tích ETABS.", "Check Model");

            dgv.DataSource = null;
            dgv.DataSource = _rows;

            if (_rows.Count > 0 && _rows.All(r => Math.Abs(r.Ptot) < 1e-9))
                Warn("Ptot vẫn bằng 0. Hãy kiểm tra Mass Summary by Story và model đã Run Analysis chưa.", "Check Model");

            btnExport.Enabled = _rows.Count > 0;
        }

        private void RunWindCheck()
        {
            const double limit = 500.0;
            if (!RequireCombo(cboWindCombo, "Chuyển vị đỉnh", "Chưa chọn tổ hợp gió.", out var windCombo)) return;

            _sap.SetPresentUnits(eUnits.kN_m_C);
            string baseStoryOverride = cboWindBaseStory != null && cboWindBaseStory.SelectedIndex >= 0
                ? cboWindBaseStory.Text
                : null;
            if (string.Equals(baseStoryOverride, "(Tự động)", StringComparison.OrdinalIgnoreCase))
                baseStoryOverride = null;

            var source = GetDisplacementSource(cboWindSource);
            var aggregationMode = GetDiaphragmAggregationMode(cboWindDiaphragmMode);
            _windRows = TopDisplacementExtractor.Calculate(_sap, windCombo, windCombo, limit,
                source, baseStoryOverride, aggregationMode);

            var displayRows = BuildWindDisplayRows(_windRows, source);

            dgvWind.DataSource = null;
            dgvWind.DataSource = displayRows;
            if (dgvWind.Columns.Contains("SourceName"))
                dgvWind.Columns["SourceName"].Visible = source == DisplacementSource.Diaphragm;

            if (_windRows.Count == 0) Warn("Không có kết quả chuyển vị đỉnh cho tổ hợp " + windCombo + ". Kiểm tra bảng Displacements và kết quả phân tích ETABS.", "Chuyển vị đỉnh");
            if (_windRows.Count > 0 && _windRows.All(r => Math.Abs(r.TopDisplacement) < 1e-12))
                Warn("Chuyển vị các tầng đang bằng 0. Hãy kiểm tra combo gió và bảng kết quả " +
                    GetDisplacementSource(cboWindSource) + " Displacements trong ETABS.\n" +
                    EtabsTableReader.LastDisplacementReadDiagnostic, "Chuyển vị đỉnh");

            btnWindExport.Enabled = _windRows.Count > 0;
        }

        private void RunWindDriftCheck()
        {
            if (!RequireCombo(cboWindDriftCombo, "Chuyển vị lệch tầng", "Chưa chọn tổ hợp gió.", out var combo)) return;
            const double limitDen = WindDriftLimitDen;

            _sap.SetPresentUnits(eUnits.kN_m_C);
            var aggregationMode = GetDiaphragmAggregationMode(cboWindDriftDiaphragmMode);
            _windDriftRows = WindDriftExtractor.Calculate(_sap, combo, combo, limitDen,
                GetDisplacementSource(cboWindDriftSource), aggregationMode);

            var displayRows = BuildWindDriftDisplayRows(_windDriftRows, limitDen);

            dgvWindDrift.DataSource = null;
            dgvWindDrift.DataSource = displayRows;

            if (_windDriftRows.Count == 0) Warn("Không có kết quả drift cho tổ hợp gió " + combo + ". Kiểm tra bảng Story Drifts và kết quả phân tích ETABS.", "Chuyển vị lệch tầng");
            if (_windDriftRows.Count > 0 && _windDriftRows.All(r => Math.Abs(r.Drift) < 1e-12))
                Warn("Drift các tầng đang bằng 0. Hãy kiểm tra tổ hợp gió và model đã Run Analysis chưa.", "Chuyển vị lệch tầng");

            btnWindDriftExport.Enabled = _windDriftRows.Count > 0;
        }

        private void RunSeismicDriftCheck()
        {
            if (!RequireCombo(cboSeisCombo, "Chuyển vị lệch tầng (động đất)", "Chưa chọn tổ hợp động đất.", out var combo)) return;
            if (!double.TryParse(txtSeisQ.Text, out var q) || q <= 0) q = 1.0;
            if (!double.TryParse(txtSeisNu.Text, out var nu) || nu <= 0) nu = 1.0;
            double limitRatio = GetSeismicLimit();

            _sap.SetPresentUnits(eUnits.kN_m_C);
            var aggregationMode = GetDiaphragmAggregationMode(cboSeisDiaphragmMode);
            _seismicDriftRows = SeismicDriftExtractor.Calculate(_sap, combo, combo, q, nu, limitRatio,
                GetDisplacementSource(cboSeisSource), aggregationMode);

            var displayRows = BuildSeismicDisplayRows(_seismicDriftRows, q, nu, limitRatio);

            dgvSeis.DataSource = null;
            dgvSeis.DataSource = displayRows;

            if (_seismicDriftRows.Count == 0) Warn("Không có kết quả drift cho tổ hợp động đất " + combo + ". Kiểm tra bảng Story Drifts và kết quả phân tích ETABS.", "Chuyển vị lệch tầng (động đất)");
            if (_seismicDriftRows.Count > 0 && _seismicDriftRows.All(r => Math.Abs(r.Drift) < 1e-12))
                Warn("Drift các tầng đang bằng 0. Hãy kiểm tra tổ hợp động đất và model đã Run Analysis chưa.", "Chuyển vị lệch tầng (động đất)");

            btnSeisExport.Enabled = _seismicDriftRows.Count > 0;
        }

        private double GetSeismicLimit()
        {
            switch (cboSeisLimit.SelectedIndex)
            {
                case 1: return 0.0075;
                case 2: return 0.010;
                default: return 0.005;
            }
        }

        private void RunAxialCheck()
        {
            if (!RequireCombo(cboAxialCombo, "Check lực dọc", "Chưa chọn combo kiểm tra.", out var combo)) return;

            double fckCube = ParseConcreteGrade(cboAxialConcrete.Text);
            if (fckCube <= 0)
            {
                Warn("Cấp bền bê tông không hợp lệ.", "Check lực dọc");
                return;
            }

            try
            {
                _sap.SetPresentUnits(eUnits.kN_m_C);
                var calc = new AxialCheckCalculator(_sap, fckCube, AxialAlphaCc, AxialGammaC, AxialColumnLimit, AxialWallLimit);
                _axialRows = calc.Build(combo);
            }
            catch (Exception ex)
            {
                Warn(ex.Message, "Check lực dọc");
                return;
            }

            dgvAxial.DataSource = null;
            dgvAxial.DataSource = _axialRows;

            int ok = _axialRows.Count(r => string.Equals(r.Result, "Thỏa mãn", StringComparison.OrdinalIgnoreCase));
            int ng = _axialRows.Count - ok;
            lblAxialInfo.Text = "Tổng: " + _axialRows.Count + "  |  Thỏa: " + ok + "  |  Không: " + ng;

            btnAxialExport.Enabled = _axialRows.Count > 0;
        }

        private static double ParseConcreteGrade(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return 0;
            var m = System.Text.RegularExpressions.Regex.Match(text, @"(\d+(?:[\.,]\d+)?)");
            if (!m.Success) return 0;
            string num = m.Groups[1].Value.Replace(',', '.');
            double v;
            return double.TryParse(num, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out v) ? v : 0;
        }

        // ---------- Gom dòng theo tầng (dùng chung cho Wind / WindDrift / Seismic) ----------

        private static List<TOut> BuildStoryRows<TIn, TOut>(
            IEnumerable<TIn> rows,
            Func<TIn, string> storyFn,
            Func<TIn, double> storyElevationFn,
            Func<TIn, string> dirFn,
            Func<TIn, double> magFn,
            Func<TIn, TIn, TOut> build) where TIn : class
        {
            var result = new List<TOut>();
            var groups = rows
                .Where(r => !EtabsHelper.IsBaseLevel(storyFn(r)))
                .GroupBy(storyFn, StringComparer.OrdinalIgnoreCase)
                .OrderByDescending(g => g.Max(storyElevationFn));

            foreach (var g in groups)
            {
                var x = g.Where(r => dirFn(r).Equals("X", StringComparison.OrdinalIgnoreCase))
                         .OrderByDescending(r => Math.Abs(magFn(r))).FirstOrDefault();
                var y = g.Where(r => dirFn(r).Equals("Y", StringComparison.OrdinalIgnoreCase))
                         .OrderByDescending(r => Math.Abs(magFn(r))).FirstOrDefault();
                if (x == null && y == null) continue;
                result.Add(build(x, y));
            }
            return result;
        }

        private List<WindGridRow> BuildWindDisplayRows(List<TopDisplacementRow> rows, DisplacementSource source)
        {
            return BuildStoryRows(rows,
                r => r.TopStory, r => r.StoryElevation, r => r.Direction, r => r.TopDisplacement,
                (x, y) =>
                {
                    var refRow = x ?? y;
                    double h = refRow.TopElevation;
                    double dx = x != null ? x.TopDisplacementMm : 0.0;
                    double dy = y != null ? y.TopDisplacementMm : 0.0;
                    double limitMm = h * 1000.0 / 500.0;
                    return new WindGridRow
                    {
                        Story = refRow.TopStory,
                        StoryElevation = refRow.StoryElevation,
                        Height = h,
                        SourceName = source == DisplacementSource.Diaphragm
                            ? (string.IsNullOrWhiteSpace(refRow.SourceTableName) ? "Diaphragm" : refRow.SourceTableName)
                            : "Story",
                        DeltaX = dx,
                        DeltaY = dy,
                        LimitMm = limitMm,
                        Check = Math.Max(dx, dy) <= limitMm ? "OK" : "NG"
                    };
                });
        }

        private List<WindDriftGridRow> BuildWindDriftDisplayRows(List<WindDriftRow> rows, double limitDen)
        {
            double limit = limitDen > 0 ? 1.0 / limitDen : 0.0;
            return BuildStoryRows(rows,
                r => r.Story, r => r.Elevation, r => r.Direction, r => r.Drift,
                (x, y) =>
                {
                    var refRow = x ?? y;
                    double driftX = x != null ? x.Drift : 0.0;
                    double driftY = y != null ? y.Drift : 0.0;
                    return new WindDriftGridRow
                    {
                        Story = refRow.Story,
                        Elevation = refRow.Elevation,
                        Height = refRow.Height,
                        DriftX = driftX,
                        DriftY = driftY,
                        Limit = limit,
                        Check = Math.Max(driftX, driftY) <= limit ? "OK" : "NG"
                    };
                });
        }

        private List<SeisGridRow> BuildSeismicDisplayRows(List<SeismicDriftRow> rows, double q, double nu, double limit)
        {
            double allow = (q * nu) > 0 ? limit / (q * nu) : 0.0;
            return BuildStoryRows(rows,
                r => r.Story, r => r.Elevation, r => r.Direction, r => r.Drift,
                (x, y) =>
                {
                    var refRow = x ?? y;
                    double driftX = x != null ? x.Drift : 0.0;
                    double driftY = y != null ? y.Drift : 0.0;
                    double driftMax = Math.Max(driftX, driftY);
                    return new SeisGridRow
                    {
                        Story = refRow.Story,
                        Elevation = refRow.Elevation,
                        Height = refRow.Height,
                        DriftX = driftX,
                        DriftY = driftY,
                        DriftMax = driftMax,
                        AllowLimit = allow,
                        Check = allow > 0 && driftMax <= allow ? "OK" : "NG"
                    };
                });
        }

        // ---------- Lớp dữ liệu hiển thị ----------

        private class WindGridRow
        {
            public string Story { get; set; }
            public double StoryElevation { get; set; }
            public double Height { get; set; }
            public string SourceName { get; set; }
            public double DeltaX { get; set; }
            public double DeltaY { get; set; }
            public double LimitMm { get; set; }
            public string Check { get; set; }
        }

        private class WindDriftGridRow
        {
            public string Story { get; set; }
            public double Elevation { get; set; }
            public double Height { get; set; }
            public double DriftX { get; set; }
            public double DriftY { get; set; }
            public double Limit { get; set; }
            public string Check { get; set; }
        }

        private class SeisGridRow
        {
            public string Story { get; set; }
            public double Elevation { get; set; }
            public double Height { get; set; }
            public double DriftX { get; set; }
            public double DriftY { get; set; }
            public double DriftMax { get; set; }
            public double AllowLimit { get; set; }
            public string Check { get; set; }
        }
    }
}
