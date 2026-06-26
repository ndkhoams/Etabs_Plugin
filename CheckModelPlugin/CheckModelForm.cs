using ETABSv1;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace CheckModelPlugin
{
    public class PDeltaCheckForm : Form
    {
        private const int CtrlHeight = 26;
        private const int BarRowHeight = 34;

        private readonly cSapModel _sap;

        private ComboBox cboCombo;
        private TextBox txtQ;
        private Button btnRun, btnExport, btnClose;
        private DataGridView dgv;
        private List<PDeltaCheckRow> _rows = new List<PDeltaCheckRow>();
        private double _qFactor = 1.0;

        private ComboBox cboWindCombo;
        private Button btnWindRun, btnWindExport;
        private DataGridView dgvWind;
        private List<TopDisplacementRow> _windRows = new List<TopDisplacementRow>();

        private ComboBox cboWindDriftCombo;
        private Button btnWindDriftRun, btnWindDriftExport;
        private DataGridView dgvWindDrift;
        private List<WindDriftRow> _windDriftRows = new List<WindDriftRow>();

        private ComboBox cboSeisCombo, cboSeisLimit;
        private TextBox txtSeisQ, txtSeisNu;
        private Button btnSeisRun, btnSeisExport;
        private DataGridView dgvSeis;
        private List<SeismicDriftRow> _seismicDriftRows = new List<SeismicDriftRow>();

        private const double WindDriftLimitDen = 500.0;

        public PDeltaCheckForm(cSapModel sap)
        {
            _sap = sap;
            InitializeComponent();
            LoadCombos();
        }

        private void InitializeComponent()
        {
            Text = "Check Model";
            Width = 1320;
            Height = 780;
            MinimumSize = new Size(1180, 700);
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Arial", 9F);

            var tabs = new TabControl { Dock = DockStyle.Fill, Font = new Font("Arial", 9F) };
            Controls.Add(tabs);

            var tabPDelta = new TabPage("Check P-Delta");
            var tabWind = new TabPage("Chuyển vị đỉnh do gió");
            var tabWindDrift = new TabPage("Chuyển vị lệch tầng do gió");
            var tabSeis = new TabPage("Chuyển vị lệch tầng do động đất");
            tabs.TabPages.Add(tabPDelta);
            tabs.TabPages.Add(tabWind);
            tabs.TabPages.Add(tabWindDrift);
            tabs.TabPages.Add(tabSeis);

            BuildPDeltaTab(tabPDelta);
            BuildWindTab(tabWind);
            BuildWindDriftTab(tabWindDrift);
            BuildSeismicDriftTab(tabSeis);
        }

        // ---------- Scaffold dùng chung cho mọi tab (bảo đảm căn hàng đồng nhất) ----------

        private DataGridView BuildScaffold(TabPage tab, string title, string standard,
            string condition, string note, out FlowLayoutPanel bar)
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 6, Padding = new Padding(12)
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));   // title
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));   // subtitle
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));   // condition
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 58));   // groupbox (thanh nhập + nút)
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 56));   // diễn giải (xuống dòng riêng)
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));   // lưới kết quả
            tab.Controls.Add(root);

            root.Controls.Add(MakeTitle(title), 0, 0);
            root.Controls.Add(MakeSubtitle(standard), 0, 1);
            root.Controls.Add(MakeCondition(condition), 0, 2);

            var box = new GroupBox
            {
                Dock = DockStyle.Fill, Text = "Tổ hợp kiểm tra", Padding = new Padding(10, 4, 10, 6)
            };
            root.Controls.Add(box, 0, 3);

            bar = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false, Margin = new Padding(0)
            };
            box.Controls.Add(bar);

            // Diễn giải nằm thành một hàng riêng ngay dưới hàng nút, full width, tự động xuống dòng.
            root.Controls.Add(MakeNote(note), 0, 4);

            var grid = CreateGrid();
            root.Controls.Add(grid, 0, 5);
            return grid;
        }

        // ---------- Các tab ----------

        private void BuildPDeltaTab(TabPage tab)
        {
            dgv = BuildScaffold(tab,
                "KIỂM TRA ĐIỀU KIỆN P-DELTA",
                "(Theo TCVN 9386-1:2025)",
                "θ = q × drift × Ptot / Vtot  |  drift = Δ/h từ ETABS Story Drift",
                "Tổ hợp dùng chung để lấy drift = Δ/h và Vtot theo cả hai phương X, Y. Ptot tự động lấy từ Mass Summary by Story: Mass × 9.80665 và cộng dồn từ mái xuống.",
                out var bar);

            bar.Controls.Add(MakeFieldLabel("Tổ hợp:", 68));
            cboCombo = MakeCombo(240); bar.Controls.Add(cboCombo);
            bar.Controls.Add(MakeFieldLabel("q:", 22));
            txtQ = MakeTextBox("1.5", 60); bar.Controls.Add(txtQ);

            btnRun = MakeButton("Tính kiểm tra"); btnRun.Click += (s, e) => RunCheck(); bar.Controls.Add(btnRun);
            btnExport = MakeButton("Xuất Excel"); btnExport.Enabled = false; btnExport.Click += (s, e) => ExportExcel(); bar.Controls.Add(btnExport);
            btnClose = MakeButton("Đóng"); btnClose.Width = 84; btnClose.Click += (s, e) => Close(); bar.Controls.Add(btnClose);

            AddPDeltaGridColumns();
        }

        private void BuildWindTab(TabPage tab)
        {
            dgvWind = BuildScaffold(tab,
                "KIỂM TRA CHUYỂN VỊ ĐỈNH CÔNG TRÌNH",
                "(Theo TCVN 2737:2023)",
                "Điều kiện kiểm tra: f ≤ fu  |  Chuyển vị ngang tổng thể giới hạn H/500",
                "Giới hạn chuyển vị ngang tổng thể mặc định là H/500. H được tính là khoảng cách từ mặt móng đến trục của xà đỡ mái.",
                out var bar);

            bar.Controls.Add(MakeFieldLabel("Tổ hợp gió:", 78));
            cboWindCombo = MakeCombo(240); bar.Controls.Add(cboWindCombo);

            btnWindRun = MakeButton("Tính kiểm tra"); btnWindRun.Click += (s, e) => RunWindCheck(); bar.Controls.Add(btnWindRun);
            btnWindExport = MakeButton("Xuất Excel"); btnWindExport.Enabled = false; btnWindExport.Click += (s, e) => ExportExcel(); bar.Controls.Add(btnWindExport);

            AddWindGridColumns();
        }

        private void BuildWindDriftTab(TabPage tab)
        {
            dgvWindDrift = BuildScaffold(tab,
                "KIỂM TRA CHUYỂN VỊ LỆCH TẦNG DO TẢI TRỌNG GIÓ",
                "(Theo TCVN 2737:2023)",
                "Điều kiện: drift = Δ/h ≤ 1/500 cho từng tầng",
                "Drift lấy trực tiếp từ ETABS Story Drifts theo tổ hợp gió. Δ = drift × chiều cao tầng. Giới hạn cố định h/500 theo TCVN 2737:2023.",
                out var bar);

            bar.Controls.Add(MakeFieldLabel("Tổ hợp gió:", 78));
            cboWindDriftCombo = MakeCombo(240); bar.Controls.Add(cboWindDriftCombo);

            btnWindDriftRun = MakeButton("Tính kiểm tra"); btnWindDriftRun.Click += (s, e) => RunWindDriftCheck(); bar.Controls.Add(btnWindDriftRun);
            btnWindDriftExport = MakeButton("Xuất Excel"); btnWindDriftExport.Enabled = false; btnWindDriftExport.Click += (s, e) => ExportExcel(); bar.Controls.Add(btnWindDriftExport);

            AddWindDriftGridColumns();
        }

        private void BuildSeismicDriftTab(TabPage tab)
        {
            dgvSeis = BuildScaffold(tab,
                "KIỂM TRA CHUYỂN VỊ LỆCH TẦNG DO TẢI TRỌNG ĐỘNG ĐẤT",
                "(Theo TCVN 9386-1:2025)",
                "Điều kiện hạn chế hư hỏng: dr·ν ≤ limit·h  ⇔  drift ≤ limit/(ν·q)",
                "drift = de/h (đàn hồi) lấy từ ETABS Story Drifts. Tổ hợp drift là động đất thuần theo quy tắc phương 1.0EX + 0.3EY (KHÔNG dùng tổ hợp trọng lực G+Q+E — tổ hợp đó chỉ dùng cho nội lực & P-Delta). dr = q × de là chuyển vị lệch tầng thiết kế. ν: hệ số chiết giảm (0.4 – 0.5). limit: 0.005 (giòn) / 0.0075 (dẻo) / 0.010 (không cản trở).",
                out var bar);

            bar.Controls.Add(MakeFieldLabel("Tổ hợp động đất:", 110));
            cboSeisCombo = MakeCombo(220); bar.Controls.Add(cboSeisCombo);
            bar.Controls.Add(MakeFieldLabel("q:", 22));
            txtSeisQ = MakeTextBox("1.5", 50); bar.Controls.Add(txtSeisQ);
            bar.Controls.Add(MakeFieldLabel("ν:", 22));
            txtSeisNu = MakeTextBox("0.5", 50); bar.Controls.Add(txtSeisNu);
            bar.Controls.Add(MakeFieldLabel("limit:", 38));
            cboSeisLimit = MakeCombo(150);
            cboSeisLimit.Items.AddRange(new object[] { "0.005 (giòn)", "0.0075 (dẻo)", "0.010 (không cản trở)" });
            cboSeisLimit.SelectedIndex = 0;
            bar.Controls.Add(cboSeisLimit);

            btnSeisRun = MakeButton("Tính kiểm tra"); btnSeisRun.Click += (s, e) => RunSeismicDriftCheck(); bar.Controls.Add(btnSeisRun);
            btnSeisExport = MakeButton("Xuất Excel"); btnSeisExport.Enabled = false; btnSeisExport.Click += (s, e) => ExportExcel(); bar.Controls.Add(btnSeisExport);

            AddSeismicDriftGridColumns();
        }

        // ---------- Factory tạo control căn chỉnh đồng nhất ----------

        private static Label MakeTitle(string text) => new Label
        {
            Text = text, Dock = DockStyle.Fill, Font = new Font("Arial", 14F, FontStyle.Bold),
            TextAlign = ContentAlignment.MiddleCenter
        };

        private static Label MakeSubtitle(string text) => new Label
        {
            Text = text, Dock = DockStyle.Fill, Font = new Font("Arial", 10F, FontStyle.Italic),
            TextAlign = ContentAlignment.MiddleCenter
        };

        private static Label MakeCondition(string text) => new Label
        {
            Text = text, Dock = DockStyle.Fill, Font = new Font("Arial", 10F),
            ForeColor = Color.DarkBlue, TextAlign = ContentAlignment.MiddleCenter
        };

        private static Label MakeNote(string text) => new Label
        {
            Text = text, Dock = DockStyle.Fill, AutoSize = false, Font = new Font("Arial", 8.5F),
            ForeColor = Color.DimGray, TextAlign = ContentAlignment.TopLeft,
            Padding = new Padding(2, 2, 2, 0)
        };

        private static Label MakeFieldLabel(string text, int width) => new Label
        {
            Text = text, AutoSize = false, Width = width, Height = CtrlHeight,
            TextAlign = ContentAlignment.MiddleLeft, Margin = new Padding(0, 4, 10, 0)
        };

        private static ComboBox MakeCombo(int width) => new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList, Width = width,
            Margin = new Padding(0, 5, 26, 0)
        };

        private static TextBox MakeTextBox(string value, int width) => new TextBox
        {
            Text = value, Width = width, Margin = new Padding(0, 6, 26, 0)
        };

        private static Button MakeButton(string text) => new Button
        {
            Text = text, Width = 118, Height = CtrlHeight, Margin = new Padding(0, 4, 14, 0)
        };

        private DataGridView CreateGrid()
        {
            return new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                AllowUserToAddRows = false,
                ReadOnly = true,
                BackgroundColor = SystemColors.ControlLightLight,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(0, 8, 0, 0)
            };
        }

        private void AddColumn(DataGridView grid, string property, string header, int width, string format = null, bool fill = false)
        {
            var col = new DataGridViewTextBoxColumn
            {
                DataPropertyName = property,
                HeaderText = header,
                Width = width,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                AutoSizeMode = fill ? DataGridViewAutoSizeColumnMode.Fill : DataGridViewAutoSizeColumnMode.None
            };
            if (!string.IsNullOrWhiteSpace(format)) col.DefaultCellStyle.Format = format;
            grid.Columns.Add(col);
        }

        private void AddPDeltaGridColumns()
        {
            dgv.Columns.Clear();
            AddColumn(dgv, "Direction", "Phương", 60);
            AddColumn(dgv, "Story", "Tầng", 80);
            AddColumn(dgv, "ElasticDrift", "drift", 120, "N5");
            AddColumn(dgv, "DesignDrift", "q × drift", 125, "N5");
            AddColumn(dgv, "Ptot", "Ptot (kN)", 105, "0");
            AddColumn(dgv, "Vtot", "Vtot (kN)", 105, "0");
            AddColumn(dgv, "Theta", "θ", 110, "N3");
            AddColumn(dgv, "Amplification", "1/(1-θ)", 90, "N3");
            AddColumn(dgv, "Conclusion", "Kết luận", 300, null, true);
        }

        private void AddWindGridColumns()
        {
            dgvWind.Columns.Clear();
            AddColumn(dgvWind, "Story", "Tầng", 90);
            AddColumn(dgvWind, "StoryElevation", "Cao độ tầng (m)", 120, "+0.000;-0.000;0.000");
            AddColumn(dgvWind, "Height", "H (m)", 110, "N3");
            AddColumn(dgvWind, "DeltaX", "ΔX (mm)", 120, "N1");
            AddColumn(dgvWind, "DeltaY", "ΔY (mm)", 120, "N1");
            AddColumn(dgvWind, "LimitMm", "H/500 (mm)", 120, "N0");
            AddColumn(dgvWind, "Check", "Kiểm tra", 250, null, true);
        }

        private void AddWindDriftGridColumns()
        {
            dgvWindDrift.Columns.Clear();
            AddColumn(dgvWindDrift, "Story", "Tầng", 90);
            AddColumn(dgvWindDrift, "Elevation", "Cao độ (m)", 100, "+0.000;-0.000;0.000");
            AddColumn(dgvWindDrift, "Height", "h tầng (m)", 90, "N3");
            AddColumn(dgvWindDrift, "DriftX", "drift X", 110, "0.000000");
            AddColumn(dgvWindDrift, "DriftY", "drift Y", 110, "0.000000");
            AddColumn(dgvWindDrift, "Limit", "Giới hạn 1/500", 120, "0.000000");
            AddColumn(dgvWindDrift, "Check", "Kiểm tra", 200, null, true);
        }

        private void AddSeismicDriftGridColumns()
        {
            dgvSeis.Columns.Clear();
            AddColumn(dgvSeis, "Story", "Tầng", 90);
            AddColumn(dgvSeis, "Elevation", "Cao độ (m)", 95, "+0.000;-0.000;0.000");
            AddColumn(dgvSeis, "Height", "h tầng (m)", 85, "N3");
            AddColumn(dgvSeis, "DriftX", "drift X (de/h)", 105, "0.000000");
            AddColumn(dgvSeis, "DriftY", "drift Y (de/h)", 105, "0.000000");
            AddColumn(dgvSeis, "DriftMax", "drift max", 100, "0.000000");
            AddColumn(dgvSeis, "AllowLimit", "Giới hạn limit/(ν·q)", 135, "0.000000");
            AddColumn(dgvSeis, "Check", "Kiểm tra", 150, null, true);
        }

        // ---------- Tải danh sách tổ hợp ----------

        private void LoadCombos()
        {
            var combos = PDeltaExtractor.GetLoadCombinations(_sap);
            foreach (var cbo in new[] { cboCombo, cboWindCombo, cboWindDriftCombo, cboSeisCombo })
            {
                cbo.Items.Clear();
                cbo.Items.AddRange(combos.Cast<object>().ToArray());
            }

            SelectByKeyword(cboCombo, "Vtot", "ENV_DD", "EQ", "DD", "DONGDAT", "RS", "SPEC", "E");
            SelectByKeyword(cboWindCombo, "ENV_SLS_W", "WX", "WY", "WINDX", "WINDY", "GIOX", "GIOY");
            SelectByKeyword(cboWindDriftCombo, "ENV_SLS_W", "WX", "WY", "WINDX", "WINDY", "GIOX", "GIOY");
            SelectByKeyword(cboSeisCombo, "ENV_DD", "EQ", "DDX", "DDY", "DD", "DONGDAT", "RS", "SPEC", "E");
        }

        private static void SelectByKeyword(ComboBox cbo, params string[] keys)
        {
            foreach (var key in keys)
                for (int i = 0; i < cbo.Items.Count; i++)
                    if (string.Equals(cbo.Items[i].ToString(), key, StringComparison.OrdinalIgnoreCase)) { cbo.SelectedIndex = i; return; }

            foreach (var key in keys)
                for (int i = 0; i < cbo.Items.Count; i++)
                    if (cbo.Items[i].ToString().IndexOf(key, StringComparison.OrdinalIgnoreCase) >= 0) { cbo.SelectedIndex = i; return; }

            if (cbo.Items.Count > 0 && cbo.SelectedIndex < 0) cbo.SelectedIndex = 0;
        }

        // ---------- Tính toán ----------

        private void RunCheck()
        {
            if (!double.TryParse(txtQ.Text, out var q)) q = 1.0;
            _qFactor = q;

            string combo = cboCombo.Text.Trim();
            if (string.IsNullOrWhiteSpace(combo))
            {
                MessageBox.Show("Chưa chọn tổ hợp kiểm tra.", "Check Model", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _sap.SetPresentUnits(eUnits.kN_m_C);
            _sap.Results.Setup.DeselectAllCasesAndCombosForOutput();

            _rows = new List<PDeltaCheckRow>();
            _rows.AddRange(PDeltaExtractor.Calculate(_sap, combo, combo, "X", q));
            _rows.AddRange(PDeltaExtractor.Calculate(_sap, combo, combo, "Y", q));
            _rows = _rows.OrderBy(r => r.Direction).ThenByDescending(r => r.Elevation).ToList();

            dgv.DataSource = null;
            dgv.DataSource = _rows;
            UpdateTitleSummary();

            if (_rows.Count > 0 && _rows.All(r => Math.Abs(r.Ptot) < 1e-9))
                MessageBox.Show("Ptot vẫn bằng 0. Hãy kiểm tra Mass Summary by Story và model đã Run Analysis chưa.", "Check Model", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            btnExport.Enabled = _rows.Count > 0;
        }

        private void RunWindCheck()
        {
            const double limit = 500.0;
            string windCombo = cboWindCombo.Text.Trim();

            if (string.IsNullOrWhiteSpace(windCombo))
            {
                MessageBox.Show("Chưa chọn tổ hợp gió.", "Chuyển vị đỉnh", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _sap.SetPresentUnits(eUnits.kN_m_C);
            _windRows = TopDisplacementExtractor.Calculate(_sap, windCombo, windCombo, limit);

            var displayRows = BuildWindDisplayRows(_windRows);

            dgvWind.DataSource = null;
            dgvWind.DataSource = displayRows;
            UpdateTitleSummary();

            if (_windRows.Count > 0 && _windRows.All(r => Math.Abs(r.TopDisplacement) < 1e-12))
                MessageBox.Show("Chuyển vị các tầng đang bằng 0. Hãy kiểm tra combo gió và bảng Diaphragm Center of Mass Displacements đã có dữ liệu chưa.", "Chuyển vị đỉnh", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            btnWindExport.Enabled = _windRows.Count > 0;
        }

        private void RunWindDriftCheck()
        {
            string combo = cboWindDriftCombo.Text.Trim();
            if (string.IsNullOrWhiteSpace(combo))
            {
                MessageBox.Show("Chưa chọn tổ hợp gió.", "Chuyển vị lệch tầng", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            const double limitDen = WindDriftLimitDen;

            _sap.SetPresentUnits(eUnits.kN_m_C);
            _windDriftRows = WindDriftExtractor.Calculate(_sap, combo, combo, limitDen);

            var displayRows = BuildWindDriftDisplayRows(_windDriftRows, limitDen);

            dgvWindDrift.DataSource = null;
            dgvWindDrift.DataSource = displayRows;

            if (_windDriftRows.Count > 0 && _windDriftRows.All(r => Math.Abs(r.Drift) < 1e-12))
                MessageBox.Show("Drift các tầng đang bằng 0. Hãy kiểm tra tổ hợp gió và model đã Run Analysis chưa.", "Chuyển vị lệch tầng", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            btnWindDriftExport.Enabled = _windDriftRows.Count > 0;
        }

        private void RunSeismicDriftCheck()
        {
            string combo = cboSeisCombo.Text.Trim();
            if (string.IsNullOrWhiteSpace(combo))
            {
                MessageBox.Show("Chưa chọn tổ hợp động đất.", "Chuyển vị lệch tầng (động đất)", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!double.TryParse(txtSeisQ.Text, out var q) || q <= 0) q = 1.0;
            if (!double.TryParse(txtSeisNu.Text, out var nu) || nu <= 0) nu = 1.0;
            double limitRatio = GetSeismicLimit();

            _sap.SetPresentUnits(eUnits.kN_m_C);
            _seismicDriftRows = SeismicDriftExtractor.Calculate(_sap, combo, combo, q, nu, limitRatio);

            var displayRows = BuildSeismicDisplayRows(_seismicDriftRows, q, nu, limitRatio);

            dgvSeis.DataSource = null;
            dgvSeis.DataSource = displayRows;

            if (_seismicDriftRows.Count > 0 && _seismicDriftRows.All(r => Math.Abs(r.Drift) < 1e-12))
                MessageBox.Show("Drift các tầng đang bằng 0. Hãy kiểm tra tổ hợp động đất và model đã Run Analysis chưa.", "Chuyển vị lệch tầng (động đất)", MessageBoxButtons.OK, MessageBoxIcon.Warning);

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

        private List<WindGridRow> BuildWindDisplayRows(List<TopDisplacementRow> rows)
        {
            var result = new List<WindGridRow>();
            var stories = rows
                .Where(r => !EtabsHelper.IsBaseLevel(r.TopStory))
                .GroupBy(r => r.TopStory, StringComparer.OrdinalIgnoreCase)
                .OrderByDescending(g => g.Max(x => x.TopElevation));

            foreach (var g in stories)
            {
                var x = g.Where(r => r.Direction.Equals("X", StringComparison.OrdinalIgnoreCase))
                         .OrderByDescending(r => Math.Abs(r.TopDisplacement)).FirstOrDefault();
                var y = g.Where(r => r.Direction.Equals("Y", StringComparison.OrdinalIgnoreCase))
                         .OrderByDescending(r => Math.Abs(r.TopDisplacement)).FirstOrDefault();
                var refRow = x ?? y;
                if (refRow == null) continue;

                double h = refRow.TopElevation;
                double storyElevation = refRow.StoryElevation;
                double dx = x != null ? x.TopDisplacementMm : 0.0;
                double dy = y != null ? y.TopDisplacementMm : 0.0;
                double limitMm = h * 1000.0 / 500.0;

                result.Add(new WindGridRow
                {
                    Story = refRow.TopStory,
                    StoryElevation = storyElevation,
                    Height = h,
                    DeltaX = dx,
                    DeltaY = dy,
                    LimitMm = limitMm,
                    Check = Math.Max(dx, dy) <= limitMm ? "OK" : "NG"
                });
            }

            return result;
        }

        private List<WindDriftGridRow> BuildWindDriftDisplayRows(List<WindDriftRow> rows, double limitDen)
        {
            var result = new List<WindDriftGridRow>();
            double limit = limitDen > 0 ? 1.0 / limitDen : 0.0;

            var stories = rows
                .Where(r => !EtabsHelper.IsBaseLevel(r.Story))
                .GroupBy(r => r.Story, StringComparer.OrdinalIgnoreCase)
                .OrderByDescending(g => g.Max(x => x.Elevation));

            foreach (var g in stories)
            {
                var x = g.Where(r => r.Direction.Equals("X", StringComparison.OrdinalIgnoreCase))
                         .OrderByDescending(r => Math.Abs(r.Drift)).FirstOrDefault();
                var y = g.Where(r => r.Direction.Equals("Y", StringComparison.OrdinalIgnoreCase))
                         .OrderByDescending(r => Math.Abs(r.Drift)).FirstOrDefault();
                var refRow = x ?? y;
                if (refRow == null) continue;

                double driftX = x != null ? x.Drift : 0.0;
                double driftY = y != null ? y.Drift : 0.0;

                result.Add(new WindDriftGridRow
                {
                    Story = refRow.Story,
                    Elevation = refRow.Elevation,
                    Height = refRow.Height,
                    DriftX = driftX,
                    DriftY = driftY,
                    Limit = limit,
                    Check = Math.Max(driftX, driftY) <= limit ? "OK" : "NG"
                });
            }
            return result;
        }

        private List<SeisGridRow> BuildSeismicDisplayRows(List<SeismicDriftRow> rows, double q, double nu, double limit)
        {
            var result = new List<SeisGridRow>();
            var stories = rows
                .Where(r => !EtabsHelper.IsBaseLevel(r.Story))
                .GroupBy(r => r.Story, StringComparer.OrdinalIgnoreCase)
                .OrderByDescending(g => g.Max(x => x.Elevation));

            foreach (var g in stories)
            {
                var x = g.Where(r => r.Direction.Equals("X", StringComparison.OrdinalIgnoreCase))
                         .OrderByDescending(r => Math.Abs(r.Drift)).FirstOrDefault();
                var y = g.Where(r => r.Direction.Equals("Y", StringComparison.OrdinalIgnoreCase))
                         .OrderByDescending(r => Math.Abs(r.Drift)).FirstOrDefault();
                var refRow = x ?? y;
                if (refRow == null) continue;

                double driftX = x != null ? x.Drift : 0.0;
                double driftY = y != null ? y.Drift : 0.0;
                double driftMax = Math.Max(driftX, driftY);
                double allow = (q * nu) > 0 ? limit / (q * nu) : 0.0;

                result.Add(new SeisGridRow
                {
                    Story = refRow.Story,
                    Elevation = refRow.Elevation,
                    Height = refRow.Height,
                    DriftX = driftX,
                    DriftY = driftY,
                    DriftMax = driftMax,
                    AllowLimit = allow,
                    Check = allow > 0 && driftMax <= allow ? "OK" : "NG"
                });
            }
            return result;
        }

        private class WindGridRow
        {
            public string Story { get; set; }
            public double StoryElevation { get; set; }
            public double Height { get; set; }
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

        private void UpdateTitleSummary()
        {
            double qx = _rows.Where(r => r.Direction.Equals("X", StringComparison.OrdinalIgnoreCase)).Select(r => r.Theta).DefaultIfEmpty(0).Max();
            double qy = _rows.Where(r => r.Direction.Equals("Y", StringComparison.OrdinalIgnoreCase)).Select(r => r.Theta).DefaultIfEmpty(0).Max();
            Text = string.Format("Check Model | θmax X = {0:0.0000}; θmax Y = {1:0.0000}", qx, qy);
        }

        private void ExportExcel()
        {
            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
                sfd.FileName = "Kiem_tra_chuyen_vi_TCVN.xlsx";
                if (sfd.ShowDialog() != DialogResult.OK) return;

                PDeltaExcelExporter.Export(sfd.FileName, _rows, _qFactor, _windRows, _windDriftRows, _seismicDriftRows);
                MessageBox.Show("Đã xuất: " + sfd.FileName, "Xuất Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
