using ETABSv1;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Etabs_Ultimate_Tools
{
    // Tab kiểm tra phản lực cọc CÓ LỰC NGANG (nhân bản tab cọc + kiểm tra FX, FY).
    public partial class ModelCheckForm
    {
        private const string PileHTitle = "Phản lực cọc (ngang)";

        // Chỉ số cột SCT trong dgvPileHCaps: 0 = Loại cọc;
        // (kéo, nén, ngang) cho Đứng / Gió / Động đất.
        private const int CapHTensVert = 1, CapHCompVert = 2, CapHHorizVert = 3;
        private const int CapHTensWind = 4, CapHCompWind = 5, CapHHorizWind = 6;
        private const int CapHTensEq = 7, CapHCompEq = 8, CapHHorizEq = 9;

        private const int CapHNameW = 80;
        private const int CapHValW = 46;

        private ComboBox cboPileHVert, cboPileHWind, cboPileHEq;
        private DataGridView dgvPileHCaps, dgvPileHPreview;
        private Button btnPileHPreview, btnPileHExport;
        private Label lblPileHInfo;
        private List<PileReactionCase> _pileHCases = new List<PileReactionCase>();

        private void BuildPileHTab(TabPage tab)
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(12)
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            tab.Controls.Add(root);

            root.Controls.Add(MakeTitle("KIỂM TRA KHẢ NĂNG CHỊU TẢI CỦA CỌC (CÓ LỰC NGANG)"), 0, 0);
            root.Controls.Add(MakeSubtitle("(Phản lực đứng so với SCT kéo/nén; hợp lực ngang H=√(FX²+FY²) so với SCT ngang của từng tổ hợp, đơn vị kN)"), 0, 1);

            var main = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Margin = new Padding(0, 6, 0, 0)
            };
            main.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 510));
            main.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.Controls.Add(main, 0, 2);

            var left = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 5, Margin = new Padding(0, 0, 10, 0)
            };
            left.RowStyles.Add(new RowStyle(SizeType.Absolute, 118));
            left.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
            left.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            left.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            left.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            main.Controls.Add(left, 0, 0);

            var comboPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 3
            };
            comboPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160));
            comboPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (int i = 0; i < 3; i++) comboPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            left.Controls.Add(comboPanel, 0, 0);

            comboPanel.Controls.Add(MakePileLabel("Tổ hợp tải đứng:"), 0, 0);
            cboPileHVert = MakePileCombo(); comboPanel.Controls.Add(cboPileHVert, 1, 0);
            comboPanel.Controls.Add(MakePileLabel("Tổ hợp tải gió:"), 0, 1);
            cboPileHWind = MakePileCombo(); comboPanel.Controls.Add(cboPileHWind, 1, 1);
            comboPanel.Controls.Add(MakePileLabel("Tổ hợp tải động đất:"), 0, 2);
            cboPileHEq = MakePileCombo(); comboPanel.Controls.Add(cboPileHEq, 1, 2);

            left.Controls.Add(new Label
            {
                Text = "SCT kéo/nén/ngang theo loại cọc & từng tổ hợp (kN):",
                Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft
            }, 0, 1);

            var capsPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Left, ColumnCount = 1, RowCount = 2,
                Width = CapHNameW + 9 * CapHValW + 4, Margin = new Padding(0, 8, 0, 0)
            };
            capsPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
            capsPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            left.Controls.Add(capsPanel, 0, 2);

            capsPanel.Controls.Add(BuildCapsHeaderH(), 0, 0);

            dgvPileHCaps = CreateEditableGrid();
            dgvPileHCaps.ColumnHeadersVisible = false;
            dgvPileHCaps.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            dgvPileHCaps.AllowUserToResizeColumns = false;
            dgvPileHCaps.AllowUserToResizeRows = false;
            dgvPileHCaps.Margin = new Padding(0);
            capsPanel.Controls.Add(dgvPileHCaps, 0, 1);
            AddPileHCapsColumns();

            var btnRow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = false,
                Margin = new Padding(0, 6, 0, 0)
            };
            left.Controls.Add(btnRow, 0, 3);

            btnPileHPreview = new Button { Text = "Xem trước", Width = 130, Height = 38, Margin = new Padding(0, 0, 12, 0) };
            btnPileHPreview.Click += (s, e) => PreviewPileHReactions();
            btnRow.Controls.Add(btnPileHPreview);

            btnPileHExport = new Button { Text = "Xuất Excel", Width = 150, Height = 38, Enabled = false, Margin = new Padding(0, 0, 0, 0) };
            btnPileHExport.Click += (s, e) => ExportPileHReactions();
            btnRow.Controls.Add(btnPileHExport);

            lblPileHInfo = new Label
            {
                Dock = DockStyle.Fill, ForeColor = Color.DimGray, TextAlign = ContentAlignment.MiddleLeft
            };
            left.Controls.Add(lblPileHInfo, 0, 4);

            var right = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2
            };
            right.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
            right.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            main.Controls.Add(right, 1, 0);

            right.Controls.Add(new Label
            {
                Text = "Preview (gộp tất cả trường hợp tải):",
                Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft
            }, 0, 0);

            dgvPileHPreview = CreateGrid();
            right.Controls.Add(dgvPileHPreview, 0, 1);
            AddPileHGridColumns();

            lblPileHInfo.Text = "Điền SCT cọc (kéo/nén/ngang) vào bảng.";
        }

        // Header gộp 2 dòng: dòng trên = trường hợp tải, dòng dưới = Kéo/Nén/Ngang.
        private Control BuildCapsHeaderH()
        {
            var h = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 10, RowCount = 2, Margin = new Padding(0)
            };
            h.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, CapHNameW));
            for (int i = 0; i < 9; i++) h.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, CapHValW));
            h.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
            h.RowStyles.Add(new RowStyle(SizeType.Percent, 50));

            var name = MakeHeadCell("Loại cọc");
            h.Controls.Add(name, 0, 0);
            h.SetRowSpan(name, 2);

            var g1 = MakeHeadCell("Tải đứng"); h.Controls.Add(g1, 1, 0); h.SetColumnSpan(g1, 3);
            var g2 = MakeHeadCell("Tải gió"); h.Controls.Add(g2, 4, 0); h.SetColumnSpan(g2, 3);
            var g3 = MakeHeadCell("Tải động đất"); h.Controls.Add(g3, 7, 0); h.SetColumnSpan(g3, 3);

            h.Controls.Add(MakeHeadCell("Kéo"), 1, 1);
            h.Controls.Add(MakeHeadCell("Nén"), 2, 1);
            h.Controls.Add(MakeHeadCell("Ngang"), 3, 1);
            h.Controls.Add(MakeHeadCell("Kéo"), 4, 1);
            h.Controls.Add(MakeHeadCell("Nén"), 5, 1);
            h.Controls.Add(MakeHeadCell("Ngang"), 6, 1);
            h.Controls.Add(MakeHeadCell("Kéo"), 7, 1);
            h.Controls.Add(MakeHeadCell("Nén"), 8, 1);
            h.Controls.Add(MakeHeadCell("Ngang"), 9, 1);
            return h;
        }

        private void AddPileHCapsColumns()
        {
            dgvPileHCaps.Columns.Clear();
            dgvPileHCaps.Columns.Add(MakeCapCol(CapHNameW, true));
            for (int i = 0; i < 9; i++)
                dgvPileHCaps.Columns.Add(MakeCapCol(CapHValW, false));
        }

        private void AddPileHGridColumns()
        {
            dgvPileHPreview.Columns.Clear();
            dgvPileHPreview.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            AddColumn(dgvPileHPreview, "LoadType", "Trường hợp", 120, null, true);
            AddColumn(dgvPileHPreview, "PileType", "Loại cọc", 70, null, true);
            AddColumn(dgvPileHPreview, "PileId", "Số hiệu", 70, null, true);
            AddColumn(dgvPileHPreview, "Combo", "Tổ hợp", 110, null, true);
            AddColumn(dgvPileHPreview, "Reaction", "Phản lực đứng (kN)", 90, "0.#", true);
            AddColumn(dgvPileHPreview, "TensionCap", "SCT kéo", 70, "0.#", true);
            AddColumn(dgvPileHPreview, "CompressionCap", "SCT nén", 70, "0.#", true);
            AddColumn(dgvPileHPreview, "Result", "KL đứng", 75, null, true);
            AddColumn(dgvPileHPreview, "Fx", "|FX| (kN)", 70, "0.#", true);
            AddColumn(dgvPileHPreview, "Fy", "|FY| (kN)", 70, "0.#", true);
            AddColumn(dgvPileHPreview, "Horizontal", "H (kN)", 70, "0.#", true);
            AddColumn(dgvPileHPreview, "HorizontalCap", "SCT ngang", 75, "0.#", true);
            AddColumn(dgvPileHPreview, "HResult", "KL ngang", 75, null, true);
        }

        private void LoadPileHSpringTypes()
        {
            if (dgvPileHCaps == null) return;
            dgvPileHCaps.Rows.Clear();

            List<PileTypeInfo> infos = new List<PileTypeInfo>();
            try { infos = PileReactionChecker.GetPileTypeInfos(_sap); }
            catch { infos = new List<PileTypeInfo>(); }

            if (infos.Count > 0)
            {
                foreach (var info in infos)
                {
                    int idx = dgvPileHCaps.Rows.Add(info.Key, "", "", "", "", "", "", "", "", "");
                    FillRowHDefaults(dgvPileHCaps.Rows[idx], info.DefaultCap);
                }
            }
            else
            {
                foreach (var name in PileReactionChecker.GetSpringTypes(_sap))
                    dgvPileHCaps.Rows.Add(name, "", "", "", "", "", "", "", "", "");
            }

            if (dgvPileHCaps.Rows.Count == 0 && lblPileHInfo != null)
                lblPileHInfo.Text = "Model chưa khai báo loại point spring nào. Hãy gán point spring cho cọc trước.";
        }

        private void SyncPileHCapsGrid()
        {
            if (dgvPileHCaps == null) return;

            List<PileTypeInfo> infos;
            try { infos = PileReactionChecker.GetPileTypeInfos(_sap); }
            catch { return; }

            var rowByType = new Dictionary<string, DataGridViewRow>(StringComparer.OrdinalIgnoreCase);
            foreach (DataGridViewRow row in dgvPileHCaps.Rows)
            {
                if (row.IsNewRow) continue;
                string nm = Convert.ToString(row.Cells[0].Value);
                if (!string.IsNullOrWhiteSpace(nm)) rowByType[nm.Trim()] = row;
            }

            foreach (var info in infos)
            {
                DataGridViewRow row;
                if (!rowByType.TryGetValue(info.Key, out row))
                {
                    int idx = dgvPileHCaps.Rows.Add(info.Key, "", "", "", "", "", "", "", "", "");
                    row = dgvPileHCaps.Rows[idx];
                    rowByType[info.Key] = row;
                }
                FillRowHDefaults(row, info.DefaultCap);
            }
        }

        // Điền SCT tạm = Kz×0.01 vào các ô KÉO/NÉN còn trống; cột NGANG để trống cho người dùng tự nhập.
        private static void FillRowHDefaults(DataGridViewRow row, double defaultCap)
        {
            if (row == null || defaultCap <= 0) return;
            string val = Math.Round(defaultCap, 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
            int[] vertCols = { CapHTensVert, CapHCompVert, CapHTensWind, CapHCompWind, CapHTensEq, CapHCompEq };
            foreach (int c in vertCols)
            {
                object cur = row.Cells[c].Value;
                if (cur == null || string.IsNullOrWhiteSpace(cur.ToString()))
                    row.Cells[c].Value = val;
            }
        }

        private Dictionary<string, PileSpringType> ReadPileHCaps(int tensCol, int compCol, int horizCol)
        {
            var dict = new Dictionary<string, PileSpringType>(StringComparer.OrdinalIgnoreCase);
            foreach (DataGridViewRow row in dgvPileHCaps.Rows)
            {
                if (row.IsNewRow) continue;
                string name = Convert.ToString(row.Cells[0].Value);
                if (string.IsNullOrWhiteSpace(name)) continue;
                dict[name.Trim()] = new PileSpringType
                {
                    Name = name.Trim(),
                    TensionCap = ParseCap(row.Cells[tensCol].Value),
                    CompressionCap = ParseCap(row.Cells[compCol].Value),
                    HorizontalCap = ParseCap(row.Cells[horizCol].Value)
                };
            }
            return dict;
        }

        private List<PileReactionCase> BuildPileHCases()
        {
            var cases = new List<PileReactionCase>();
            AddPileHCase(cases, cboPileHVert.Text.Trim(), "TỔ HỢP TẢI ĐỨNG", "TAI DUNG",
                ReadPileHCaps(CapHTensVert, CapHCompVert, CapHHorizVert));
            AddPileHCase(cases, cboPileHWind.Text.Trim(), "TỔ HỢP TẢI GIÓ", "TAI GIO",
                ReadPileHCaps(CapHTensWind, CapHCompWind, CapHHorizWind));
            AddPileHCase(cases, cboPileHEq.Text.Trim(), "TỔ HỢP TẢI ĐỘNG ĐẤT", "TAI DONG DAT",
                ReadPileHCaps(CapHTensEq, CapHCompEq, CapHHorizEq));
            return cases;
        }

        private void AddPileHCase(List<PileReactionCase> cases, string combo, string title,
            string sheet, Dictionary<string, PileSpringType> caps)
        {
            var c = PileReactionChecker.ComputeCaseH(_sap, combo, title, sheet, caps);
            if (c != null) cases.Add(c);
        }

        private void PreviewPileHReactions()
        {
            List<PileReactionCase> cases;
            try
            {
                _sap.SetPresentUnits(eUnits.kN_m_C);
                SyncPileHCapsGrid();
                cases = BuildPileHCases();
            }
            catch (Exception ex)
            {
                Warn(ex.Message, PileHTitle);
                return;
            }

            if (cases.Count == 0)
            {
                Warn("Chưa chọn tổ hợp tải nào (cần chọn ít nhất 1 trong 3 trường hợp).", PileHTitle);
                return;
            }

            _pileHCases = cases;

            var preview = new List<PilePreviewRowH>();
            foreach (var c in cases)
                foreach (var row in c.Rows)
                    preview.Add(new PilePreviewRowH
                    {
                        LoadType = c.Title,
                        PileType = row.PileType,
                        PileId = row.PileId,
                        Combo = row.Combo,
                        Reaction = row.Reaction,
                        TensionCap = row.TensionCap,
                        CompressionCap = row.CompressionCap,
                        Result = row.Result,
                        Fx = row.Fx,
                        Fy = row.Fy,
                        Horizontal = row.Horizontal,
                        HorizontalCap = row.HorizontalCap,
                        HResult = row.HResult
                    });

            dgvPileHPreview.DataSource = null;
            dgvPileHPreview.DataSource = preview;

            int fail = preview.Count(p =>
                (!string.IsNullOrEmpty(p.Result) && p.Result.IndexOf("Không", StringComparison.OrdinalIgnoreCase) >= 0)
                || (!string.IsNullOrEmpty(p.HResult) && p.HResult.IndexOf("Không", StringComparison.OrdinalIgnoreCase) >= 0));

            lblPileHInfo.Text = "Số trường hợp: " + cases.Count + "  |  Tổng dòng: " + preview.Count + "  |  Không đạt: " + fail;

            btnPileHExport.Enabled = preview.Count > 0;

            if (preview.Count == 0)
                Warn("Không đọc được phản lực cọc. Kiểm tra: (1) các điểm cọc đã gán point spring chưa, (2) model đã Run Analysis chưa, (3) tổ hợp tải đã chọn đúng chưa.", PileHTitle);
        }

        private void ExportPileHReactions()
        {
            if (_pileHCases == null || _pileHCases.Count == 0 || _pileHCases.All(c => c.Rows == null || c.Rows.Count == 0))
            {
                Warn("Chưa có dữ liệu. Hãy bấm Xem trước trước khi xuất.", PileHTitle);
                return;
            }

            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
                sfd.FileName = "KiemTra_PhanLucCoc_Ngang.xlsx";
                if (sfd.ShowDialog() != DialogResult.OK) return;

                try
                {
                    PileReactionExporterH.Export(sfd.FileName, _pileHCases);
                }
                catch (Exception ex)
                {
                    Warn(ex.Message, PileHTitle);
                    return;
                }

                Info("Đã xuất: " + sfd.FileName, PileHTitle);
            }
        }

        private class PilePreviewRowH
        {
            public string LoadType { get; set; }
            public string PileType { get; set; }
            public string PileId { get; set; }
            public string Combo { get; set; }
            public double Reaction { get; set; }
            public double TensionCap { get; set; }
            public double CompressionCap { get; set; }
            public string Result { get; set; }
            public double Fx { get; set; }
            public double Fy { get; set; }
            public double Horizontal { get; set; }
            public double HorizontalCap { get; set; }
            public string HResult { get; set; }
        }
    }
}
