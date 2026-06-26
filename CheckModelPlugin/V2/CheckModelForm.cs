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
        private readonly cSapModel _sap;

        private ComboBox cboComboX, cboComboY;
        private TextBox txtQ;
        private Button btnRun, btnExport, btnClose;
        private DataGridView dgv;
        private List<PDeltaCheckRow> _rows = new List<PDeltaCheckRow>();
        private double _qFactor = 1.0;

        private ComboBox cboWindCombo;
        private Button btnWindRun, btnWindExport;
        private DataGridView dgvWind;
        private List<TopDisplacementRow> _windRows = new List<TopDisplacementRow>();

        public PDeltaCheckForm(cSapModel sap)
        {
            _sap = sap;
            InitializeComponent();
            LoadCombos();
        }

        private void InitializeComponent()
        {
            Text = "Check Model";
            Width = 1240;
            Height = 760;
            MinimumSize = new Size(1100, 680);
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Arial", 9F);

            var tabs = new TabControl { Dock = DockStyle.Fill, Font = new Font("Arial", 9F) };
            Controls.Add(tabs);

            var tabPDelta = new TabPage("Check P-Delta");
            var tabWind = new TabPage("Chuyển vị đỉnh do gió");
            tabs.TabPages.Add(tabPDelta);
            tabs.TabPages.Add(tabWind);

            BuildPDeltaTab(tabPDelta);
            BuildWindTab(tabWind);
        }

        private void BuildPDeltaTab(TabPage tab)
        {
            var main = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 5,
                Padding = new Padding(10)
            };
            main.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            main.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
            main.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            main.RowStyles.Add(new RowStyle(SizeType.Absolute, 118));
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            tab.Controls.Add(main);

            main.Controls.Add(new Label
            {
                Text = "KIỂM TRA ĐIỀU KIỆN P-DELTA",
                Dock = DockStyle.Fill,
                Font = new Font("Arial", 14F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter
            }, 0, 0);

            main.Controls.Add(new Label
            {
                Text = "(Theo TCVN 9386-1:2025)",
                Dock = DockStyle.Fill,
                Font = new Font("Arial", 10F, FontStyle.Italic),
                TextAlign = ContentAlignment.MiddleCenter
            }, 0, 1);

            main.Controls.Add(new Label
            {
                Text = "θ = q × drift × Ptot / Vtot | drift = Δ/h từ ETABS Story Drift",
                Dock = DockStyle.Fill,
                Font = new Font("Arial", 10F),
                ForeColor = Color.DarkBlue,
                TextAlign = ContentAlignment.MiddleCenter
            }, 0, 2);

            var box = new GroupBox { Dock = DockStyle.Fill, Text = "Tổ hợp kiểm tra", Padding = new Padding(12) };
            main.Controls.Add(box, 0, 3);

            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 7, RowCount = 3 };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 230));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 230));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 35));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            box.Controls.Add(layout);

            AddLabelCell(layout, "Combo X:", 0, 0);
            cboComboX = AddComboCell(layout, 1, 0);
            AddLabelCell(layout, "Combo Y:", 2, 0);
            cboComboY = AddComboCell(layout, 3, 0);
            AddLabelCell(layout, "q:", 4, 0);
            txtQ = new TextBox { Text = "1.5", Width = 70, Anchor = AnchorStyles.Left | AnchorStyles.Top };
            layout.Controls.Add(txtQ, 5, 0);

            AddLabelCell(layout, "Ptot:", 0, 1);
            var lblPtotAuto = new Label
            {
                Text = "Tự động lấy từ Mass Summary by Story",
                Dock = DockStyle.Fill,
                ForeColor = Color.Green,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 0, 16, 0)
            };
            layout.SetColumnSpan(lblPtotAuto, 3);
            layout.Controls.Add(lblPtotAuto, 1, 1);

            var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = false, Margin = new Padding(0) };
            layout.SetColumnSpan(buttons, 3);
            layout.Controls.Add(buttons, 4, 1);

            btnRun = new Button { Text = "Tính kiểm tra", Width = 120, Height = 30, Margin = new Padding(0, 0, 10, 0) };
            btnRun.Click += (s, e) => RunCheck();
            buttons.Controls.Add(btnRun);

            btnExport = new Button { Text = "Xuất Excel", Width = 120, Height = 30, Enabled = false, Margin = new Padding(0, 0, 10, 0) };
            btnExport.Click += (s, e) => ExportExcel();
            buttons.Controls.Add(btnExport);

            btnClose = new Button { Text = "Đóng", Width = 90, Height = 30, Margin = new Padding(0) };
            btnClose.Click += (s, e) => Close();
            buttons.Controls.Add(btnClose);

            var note = new Label
            {
                Dock = DockStyle.Fill,
                Text = "Combo X/Y dùng chung để lấy drift = Δ/h và Vtot theo phương tương ứng. Ptot tự động lấy từ Mass Summary by Story: Mass × 9.80665 và cộng dồn từ mái xuống.",
                ForeColor = Color.DimGray,
                TextAlign = ContentAlignment.MiddleLeft
            };
            layout.SetColumnSpan(note, 5);
            layout.Controls.Add(note, 0, 2);

            dgv = CreateGrid();
            AddPDeltaGridColumns();
            main.Controls.Add(dgv, 0, 4);
        }

        private void BuildWindTab(TabPage tab)
        {
            var main = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 5,
                Padding = new Padding(10)
            };
            main.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            main.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
            main.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            main.RowStyles.Add(new RowStyle(SizeType.Absolute, 105));
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            tab.Controls.Add(main);

            main.Controls.Add(new Label
            {
                Text = "KIỂM TRA CHUYỂN VỊ ĐỈNH CÔNG TRÌNH",
                Dock = DockStyle.Fill,
                Font = new Font("Arial", 14F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter
            }, 0, 0);

            main.Controls.Add(new Label
            {
                Text = "(Theo TCVN 2737:2023)",
                Dock = DockStyle.Fill,
                Font = new Font("Arial", 10F, FontStyle.Italic),
                TextAlign = ContentAlignment.MiddleCenter
            }, 0, 1);

            main.Controls.Add(new Label
            {
                Text = "Điều kiện kiểm tra: f ≤ fu | Chuyển vị ngang tổng thể giới hạn H/500",
                Dock = DockStyle.Fill,
                Font = new Font("Arial", 10F),
                ForeColor = Color.DarkBlue,
                TextAlign = ContentAlignment.MiddleCenter
            }, 0, 2);

            var box = new GroupBox { Dock = DockStyle.Fill, Text = "Tổ hợp kiểm tra", Padding = new Padding(12) };
            main.Controls.Add(box, 0, 3);

            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 5, RowCount = 2 };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 230));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 230));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 250));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            box.Controls.Add(layout);

            AddLabelCell(layout, "Tổ hợp gió:", 0, 0);
            cboWindCombo = AddComboCell(layout, 1, 0);

            var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = false, Margin = new Padding(0) };
            layout.SetColumnSpan(buttons, 2);
            layout.Controls.Add(buttons, 2, 0);

            btnWindRun = new Button { Text = "Tính kiểm tra", Width = 120, Height = 30, Margin = new Padding(0, 0, 10, 0) };
            btnWindRun.Click += (s, e) => RunWindCheck();
            buttons.Controls.Add(btnWindRun);

            btnWindExport = new Button { Text = "Xuất Excel", Width = 120, Height = 30, Enabled = false, Margin = new Padding(0, 0, 10, 0) };
            btnWindExport.Click += (s, e) => ExportExcel();
            buttons.Controls.Add(btnWindExport);

            var note = new Label
            {
                Dock = DockStyle.Fill,
                Text = "Giới hạn chuyển vị ngang tổng thể mặc định là H/500. H được tính là khoảng cách từ mặt móng đến trục của xà đỡ mái.",
                ForeColor = Color.DimGray,
                TextAlign = ContentAlignment.MiddleLeft
            };
            layout.SetColumnSpan(note, 5);
            layout.Controls.Add(note, 0, 1);

            dgvWind = CreateGrid();
            AddWindGridColumns();
            main.Controls.Add(dgvWind, 0, 4);
        }

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
                BorderStyle = BorderStyle.FixedSingle
            };
        }

        private static void AddLabelCell(TableLayoutPanel layout, string text, int col, int row)
        {
            layout.Controls.Add(new Label { Text = text, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Margin = new Padding(0, 0, 4, 0) }, col, row);
        }

        private static ComboBox AddComboCell(TableLayoutPanel layout, int col, int row)
        {
            var cb = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(0, 3, 16, 3) };
            layout.Controls.Add(cb, col, row);
            return cb;
        }

        private void AddPDeltaGridColumns()
        {
            dgv.Columns.Clear();
            AddColumn(dgv, "Direction", "Phương", 55);
            AddColumn(dgv, "Story", "Tầng", 75);
            AddColumn(dgv, "ElasticDrift", "drift", 120, "N5");
            AddColumn(dgv, "DesignDrift", "q × drift", 125, "N5");
            AddColumn(dgv, "Ptot", "Ptot (kN)", 105, "0");
            AddColumn(dgv, "Vtot", "Vtot (kN)", 105, "0");
            AddColumn(dgv, "Theta", "θ", 115, "N3");
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

        private void LoadCombos()
        {
            var combos = PDeltaExtractor.GetLoadCombinations(_sap);
            foreach (var cbo in new[] { cboComboX, cboComboY, cboWindCombo })
            {
                cbo.Items.Clear();
                cbo.Items.AddRange(combos.Cast<object>().ToArray());
            }

            SelectByKeyword(cboComboX, "Vtot", "EX", "X");
            SelectByKeyword(cboComboY, "Vtot", "EY", "Y");
            SelectByKeyword(cboWindCombo, "ENV_SLS_W", "WX", "WY", "WINDX", "WINDY", "GIOX", "GIOY");
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

        private void RunCheck()
        {
            if (!double.TryParse(txtQ.Text, out var q)) q = 1.0;
            _qFactor = q;

            string comboX = cboComboX.Text.Trim();
            string comboY = cboComboY.Text.Trim();
            if (string.IsNullOrWhiteSpace(comboX) || string.IsNullOrWhiteSpace(comboY))
            {
                MessageBox.Show("Chưa chọn đủ Combo X/Y.", "Check Model", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _sap.SetPresentUnits(eUnits.kN_m_C);
            _sap.Results.Setup.DeselectAllCasesAndCombosForOutput();

            _rows = new List<PDeltaCheckRow>();
            _rows.AddRange(PDeltaExtractor.Calculate(_sap, comboX, comboX, "X", q));
            _rows.AddRange(PDeltaExtractor.Calculate(_sap, comboY, comboY, "Y", q));
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

        // FIX: lấy dòng có |chuyển vị| lớn nhất cho mỗi phương để đồng nhất với
        // PDeltaExcelExporter (trước đây dùng FirstOrDefault gây lệch số liệu).
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

                PDeltaExcelExporter.Export(sfd.FileName, _rows, _qFactor, _windRows);
                MessageBox.Show("Đã xuất: " + sfd.FileName, "Xuất Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}