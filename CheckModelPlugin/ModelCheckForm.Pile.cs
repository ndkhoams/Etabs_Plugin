using ETABSv1;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Etabs_Ultimate_Tools
{
    // Tab kiểm tra phản lực cọc (phần partial của ModelCheckForm).
    public partial class ModelCheckForm
    {
        private const string PileTitle = "Phản lực cọc";

        // Chỉ số cột SCT theo từng trường hợp tải trong dgvPileCaps.
        // 0 = Loại cọc; (kéo, nén) cho Đứng / Gió / Động đất.
        private const int CapTensVert = 1, CapCompVert = 2;
        private const int CapTensWind = 3, CapCompWind = 4;
        private const int CapTensEq = 5, CapCompEq = 6;

        // Bề rộng cố định các cột bảng SCT (giữ bảng hẹp, nhường chỗ cho preview).
        private const int CapNameW = 86;
        private const int CapValW = 58;

        private void BuildPileTab(TabPage tab)
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(12)
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            tab.Controls.Add(root);

            root.Controls.Add(MakeTitle("KIỂM TRA KHẢ NĂNG CHỊU TẢI CỦA CỌC"), 0, 0);
            root.Controls.Add(MakeSubtitle("(So sánh phản lực đầu cọc theo phương đứng với SCT chịu kéo/nén của từng tổ hợp, đơn vị kN)"), 0, 1);

            var main = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Margin = new Padding(0, 6, 0, 0)
            };
            main.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 470));
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
            cboPileVert = MakePileCombo(); comboPanel.Controls.Add(cboPileVert, 1, 0);
            comboPanel.Controls.Add(MakePileLabel("Tổ hợp tải gió:"), 0, 1);
            cboPileWind = MakePileCombo(); comboPanel.Controls.Add(cboPileWind, 1, 1);
            comboPanel.Controls.Add(MakePileLabel("Tổ hợp tải động đất:"), 0, 2);
            cboPileEq = MakePileCombo(); comboPanel.Controls.Add(cboPileEq, 1, 2);

            left.Controls.Add(new Label
            {
                Text = "SCT chịu kéo/nén theo loại cọc & từng tổ hợp (kN) — tự điền Kz×0.01:",
                Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft
            }, 0, 1);

            // Khu vực nhập SCT = header gộp 2 dòng + grid (cùng bề rộng cột cố định).
            var capsPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Left, ColumnCount = 1, RowCount = 2,
                Width = CapNameW + 6 * CapValW + 4, Margin = new Padding(0, 8, 0, 0)
            };
            capsPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
            capsPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            left.Controls.Add(capsPanel, 0, 2);

            capsPanel.Controls.Add(BuildCapsHeader(), 0, 0);

            dgvPileCaps = CreateEditableGrid();
            dgvPileCaps.ColumnHeadersVisible = false;
            dgvPileCaps.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            dgvPileCaps.AllowUserToResizeColumns = false;
            dgvPileCaps.AllowUserToResizeRows = false;
            dgvPileCaps.Margin = new Padding(0);
            capsPanel.Controls.Add(dgvPileCaps, 0, 1);
            AddPileCapsColumns();

            var btnRow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = false,
                Margin = new Padding(0, 6, 0, 0)
            };
            left.Controls.Add(btnRow, 0, 3);

            btnPilePreview = new Button { Text = "Xem trước", Width = 130, Height = 38, Margin = new Padding(0, 0, 12, 0) };
            btnPilePreview.Click += (s, e) => PreviewPileReactions();
            btnRow.Controls.Add(btnPilePreview);

            btnPileExport = new Button { Text = "Xuất Excel", Width = 150, Height = 38, Enabled = false, Margin = new Padding(0, 0, 0, 0) };
            btnPileExport.Click += (s, e) => ExportPileReactions();
            btnRow.Controls.Add(btnPileExport);

            lblPileInfo = new Label
            {
                Dock = DockStyle.Fill, ForeColor = Color.DimGray, TextAlign = ContentAlignment.MiddleLeft
            };
            left.Controls.Add(lblPileInfo, 0, 4);

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

            dgvPilePreview = CreateGrid();
            right.Controls.Add(dgvPilePreview, 0, 1);
            AddPileGridColumns();

            lblPileInfo.Text = "SCT tạm được tự điền = Kz×0.01; chỉnh lại nếu cần, chọn 3 tổ hợp rồi bấm Xem trước.";
        }

        // Header gộp 2 dòng: dòng trên = trường hợp tải, dòng dưới = Kéo/Nén.
        private Control BuildCapsHeader()
        {
            var h = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 7, RowCount = 2, Margin = new Padding(0)
            };
            h.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, CapNameW));
            for (int i = 0; i < 6; i++) h.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, CapValW));
            h.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
            h.RowStyles.Add(new RowStyle(SizeType.Percent, 50));

            var name = MakeHeadCell("Loại cọc");
            h.Controls.Add(name, 0, 0);
            h.SetRowSpan(name, 2);

            var g1 = MakeHeadCell("Tải đứng"); h.Controls.Add(g1, 1, 0); h.SetColumnSpan(g1, 2);
            var g2 = MakeHeadCell("Tải gió"); h.Controls.Add(g2, 3, 0); h.SetColumnSpan(g2, 2);
            var g3 = MakeHeadCell("Tải động đất"); h.Controls.Add(g3, 5, 0); h.SetColumnSpan(g3, 2);

            h.Controls.Add(MakeHeadCell("Kéo"), 1, 1);
            h.Controls.Add(MakeHeadCell("Nén"), 2, 1);
            h.Controls.Add(MakeHeadCell("Kéo"), 3, 1);
            h.Controls.Add(MakeHeadCell("Nén"), 4, 1);
            h.Controls.Add(MakeHeadCell("Kéo"), 5, 1);
            h.Controls.Add(MakeHeadCell("Nén"), 6, 1);
            return h;
        }

        private static Label MakeHeadCell(string text) => new Label
        {
            Text = text, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter,
            BorderStyle = BorderStyle.FixedSingle, BackColor = SystemColors.ControlLight,
            Margin = new Padding(0), Font = new Font(SystemFonts.DefaultFont, FontStyle.Bold)
        };

        private static Label MakePileLabel(string text) => new Label
        {
            Text = text, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft
        };

        private static ComboBox MakePileCombo() => new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill, Margin = new Padding(0, 4, 0, 4)
        };

        private DataGridView CreateEditableGrid()
        {
            return new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                AllowUserToAddRows = false,
                BackgroundColor = SystemColors.ControlLightLight,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(0, 8, 0, 0)
            };
        }

        private void AddPileCapsColumns()
        {
            dgvPileCaps.Columns.Clear();
            dgvPileCaps.Columns.Add(MakeCapCol(CapNameW, true));
            for (int i = 0; i < 6; i++)
                dgvPileCaps.Columns.Add(MakeCapCol(CapValW, false));
        }

        private static DataGridViewTextBoxColumn MakeCapCol(int width, bool readOnly)
        {
            return new DataGridViewTextBoxColumn
            {
                Width = width, ReadOnly = readOnly,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
        }

        private void AddPileGridColumns()
        {
            dgvPilePreview.Columns.Clear();
            dgvPilePreview.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            AddColumn(dgvPilePreview, "LoadType", "Trường hợp", 150, null, true);
            AddColumn(dgvPilePreview, "PileType", "Loại cọc", 80, null, true);
            AddColumn(dgvPilePreview, "PileId", "Số hiệu cọc", 90, null, true);
            AddColumn(dgvPilePreview, "Combo", "Tổ hợp", 140, null, true);
            AddColumn(dgvPilePreview, "Reaction", "Phản lực (kN)", 100, "0.#", true);
            AddColumn(dgvPilePreview, "TensionCap", "SCT kéo (kN)", 95, "0.#", true);
            AddColumn(dgvPilePreview, "CompressionCap", "SCT nén (kN)", 95, "0.#", true);
            AddColumn(dgvPilePreview, "Result", "Kết Luận", 90, null, true);
        }

        private void LoadPileSpringTypes()
        {
            if (dgvPileCaps == null) return;
            dgvPileCaps.Rows.Clear();

            List<PileTypeInfo> infos = new List<PileTypeInfo>();
            try { infos = PileReactionChecker.GetPileTypeInfos(_sap); }
            catch { infos = new List<PileTypeInfo>(); }

            if (infos.Count > 0)
            {
                foreach (var info in infos)
                {
                    int idx = dgvPileCaps.Rows.Add(info.Key, "", "", "", "", "", "");
                    FillRowDefaults(dgvPileCaps.Rows[idx], info.DefaultCap);
                }
            }
            else
            {
                foreach (var name in PileReactionChecker.GetSpringTypes(_sap))
                    dgvPileCaps.Rows.Add(name, "", "", "", "", "", "");
            }

            if (dgvPileCaps.Rows.Count == 0 && lblPileInfo != null)
                lblPileInfo.Text = "Model chưa khai báo loại point spring nào. Hãy gán point spring cho cọc trước.";
        }

        // Đồng bộ bảng SCT với các loại cọc thực phát hiện được (thêm dòng thiếu,
        // tự điền SCT tạm = Kz×0.01 vào các ô còn trống, giữ nguyên ô đã nhập).
        private void SyncPileCapsGrid()
        {
            if (dgvPileCaps == null) return;

            List<PileTypeInfo> infos;
            try { infos = PileReactionChecker.GetPileTypeInfos(_sap); }
            catch { return; }

            var rowByType = new Dictionary<string, DataGridViewRow>(StringComparer.OrdinalIgnoreCase);
            foreach (DataGridViewRow row in dgvPileCaps.Rows)
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
                    int idx = dgvPileCaps.Rows.Add(info.Key, "", "", "", "", "", "");
                    row = dgvPileCaps.Rows[idx];
                    rowByType[info.Key] = row;
                }
                FillRowDefaults(row, info.DefaultCap);
            }
        }

        // Điền SCT tạm vào các ô kéo/nén còn trống (không ghi đè giá trị đã có).
        private static void FillRowDefaults(DataGridViewRow row, double defaultCap)
        {
            if (row == null || defaultCap <= 0) return;
            string val = Math.Round(defaultCap, 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
            for (int c = 1; c <= 6; c++)
            {
                object cur = row.Cells[c].Value;
                if (cur == null || string.IsNullOrWhiteSpace(cur.ToString()))
                    row.Cells[c].Value = val;
            }
        }

        // Đọc SCT cho 1 trường hợp tải (theo cặp cột kéo/nén tương ứng).
        private Dictionary<string, PileSpringType> ReadPileCaps(int tensCol, int compCol)
        {
            var dict = new Dictionary<string, PileSpringType>(StringComparer.OrdinalIgnoreCase);
            foreach (DataGridViewRow row in dgvPileCaps.Rows)
            {
                if (row.IsNewRow) continue;
                string name = Convert.ToString(row.Cells[0].Value);
                if (string.IsNullOrWhiteSpace(name)) continue;
                dict[name.Trim()] = new PileSpringType
                {
                    Name = name.Trim(),
                    TensionCap = ParseCap(row.Cells[tensCol].Value),
                    CompressionCap = ParseCap(row.Cells[compCol].Value)
                };
            }
            return dict;
        }

        private static double ParseCap(object value)
        {
            if (value == null) return 0.0;
            string s = value.ToString().Trim().Replace(",", ".");
            double d;
            return double.TryParse(s, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out d) ? d : 0.0;
        }

        private List<PileReactionCase> BuildPileCases()
        {
            var cases = new List<PileReactionCase>();
            AddPileCase(cases, cboPileVert.Text.Trim(), "TỔ HỢP TẢI ĐỨNG", "TAI DUNG",
                ReadPileCaps(CapTensVert, CapCompVert));
            AddPileCase(cases, cboPileWind.Text.Trim(), "TỔ HỢP TẢI GIÓ", "TAI GIO",
                ReadPileCaps(CapTensWind, CapCompWind));
            AddPileCase(cases, cboPileEq.Text.Trim(), "TỔ HỢP TẢI ĐỘNG ĐẤT", "TAI DONG DAT",
                ReadPileCaps(CapTensEq, CapCompEq));
            return cases;
        }

        private void AddPileCase(List<PileReactionCase> cases, string combo, string title,
            string sheet, Dictionary<string, PileSpringType> caps)
        {
            var c = PileReactionChecker.ComputeCase(_sap, combo, title, sheet, caps);
            if (c != null) cases.Add(c);
        }

        private void PreviewPileReactions()
        {
            List<PileReactionCase> cases;
            try
            {
                _sap.SetPresentUnits(eUnits.kN_m_C);
                SyncPileCapsGrid();
                cases = BuildPileCases();
            }
            catch (Exception ex)
            {
                Warn(ex.Message, PileTitle);
                return;
            }

            if (cases.Count == 0)
            {
                Warn("Chưa chọn tổ hợp tải nào (cần chọn ít nhất 1 trong 3 trường hợp).", PileTitle);
                return;
            }

            _pileCases = cases;

            var preview = new List<PilePreviewRow>();
            foreach (var c in cases)
                foreach (var row in c.Rows)
                    preview.Add(new PilePreviewRow
                    {
                        LoadType = c.Title,
                        PileType = row.PileType,
                        PileId = row.PileId,
                        Combo = row.Combo,
                        Reaction = row.Reaction,
                        TensionCap = row.TensionCap,
                        CompressionCap = row.CompressionCap,
                        Result = row.Result
                    });

            dgvPilePreview.DataSource = null;
            dgvPilePreview.DataSource = preview;

            int fail = preview.Count(p => !string.IsNullOrEmpty(p.Result)
                && p.Result.IndexOf("Không", StringComparison.OrdinalIgnoreCase) >= 0);

            lblPileInfo.Text = "Số trường hợp: " + cases.Count + "  |  Tổng dòng: " + preview.Count + "  |  Không đạt: " + fail;

            btnPileExport.Enabled = preview.Count > 0;

            if (preview.Count == 0)
                Warn("Không đọc được phản lực cọc. Kiểm tra: (1) các điểm cọc đã gán point spring chưa, (2) model đã Run Analysis chưa, (3) tổ hợp tải đã chọn đúng chưa.", PileTitle);
        }

        private void ExportPileReactions()
        {
            if (_pileCases == null || _pileCases.Count == 0 || _pileCases.All(c => c.Rows == null || c.Rows.Count == 0))
            {
                Warn("Chưa có dữ liệu. Hãy bấm Xem trước trước khi xuất.", PileTitle);
                return;
            }

            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
                sfd.FileName = "KiemTra_PhanLucCoc.xlsx";
                if (sfd.ShowDialog() != DialogResult.OK) return;

                try
                {
                    PileReactionExporter.Export(sfd.FileName, _pileCases);
                }
                catch (Exception ex)
                {
                    Warn(ex.Message, PileTitle);
                    return;
                }

                Info("Đã xuất: " + sfd.FileName, PileTitle);
            }
        }

        private class PilePreviewRow
        {
            public string LoadType { get; set; }
            public string PileType { get; set; }
            public string PileId { get; set; }
            public string Combo { get; set; }
            public double Reaction { get; set; }
            public double TensionCap { get; set; }
            public double CompressionCap { get; set; }
            public string Result { get; set; }
        }
    }
}
