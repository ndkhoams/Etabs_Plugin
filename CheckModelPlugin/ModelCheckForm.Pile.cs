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

            root.Controls.Add(MakeTitle("KIỂM TRA KHẢ NĂNG CHỊ TẢI CỦA CỌC"), 0, 0);
            root.Controls.Add(MakeSubtitle("(So sánh phản lực đầu cọc theo phương đứng với SCT chịu kéo/nén, đơn vị kN)"), 0, 1);

            var main = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Margin = new Padding(0, 6, 0, 0)
            };
            main.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 460));
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
                Text = "Khả năng chịu tải theo loại cọc (point spring) — nhập SCT kéo/nén:",
                Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft
            }, 0, 1);

            dgvPileCaps = CreateEditableGrid();
            left.Controls.Add(dgvPileCaps, 0, 2);
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

            btnPileExport = new Button { Text = "Xuất Excel (3 sheet)", Width = 190, Height = 38, Enabled = false, Margin = new Padding(0, 0, 0, 0) };
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
                Text = "Preview (gộp cả 3 trường hợp tải):",
                Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft
            }, 0, 0);

            dgvPilePreview = CreateGrid();
            right.Controls.Add(dgvPilePreview, 0, 1);
            AddPileGridColumns();

            lblPileInfo.Text = "Nhập SCT kéo/nén cho từng loại cọc, chọn 3 tổ hợp rồi bấm Xem trước.";
        }

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
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
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
            var c0 = new DataGridViewTextBoxColumn
            {
                HeaderText = "Loại cọc", ReadOnly = true, FillWeight = 120,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            var c1 = new DataGridViewTextBoxColumn
            {
                HeaderText = "SCT kéo (kN)", FillWeight = 100,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            var c2 = new DataGridViewTextBoxColumn
            {
                HeaderText = "SCT nén (kN)", FillWeight = 100,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvPileCaps.Columns.Add(c0);
            dgvPileCaps.Columns.Add(c1);
            dgvPileCaps.Columns.Add(c2);
        }

        private void AddPileGridColumns()
        {
            dgvPilePreview.Columns.Clear();
            dgvPilePreview.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            AddColumn(dgvPilePreview, "LoadType", "Trường hợp", 140, null, true);
            AddColumn(dgvPilePreview, "PileType", "Loại cọc", 80, null, true);
            AddColumn(dgvPilePreview, "PileId", "Số hiệu cọc", 90, null, true);
            AddColumn(dgvPilePreview, "Combo", "Tổ hợp", 120, null, true);
            AddColumn(dgvPilePreview, "Reaction", "Phản lực (kN)", 100, "0.#", true);
            AddColumn(dgvPilePreview, "TensionCap", "SCT kéo (kN)", 95, "0.#", true);
            AddColumn(dgvPilePreview, "CompressionCap", "SCT nén (kN)", 95, "0.#", true);
            AddColumn(dgvPilePreview, "Result", "Kết Luận", 90, null, true);
        }

        private void LoadPileSpringTypes()
        {
            if (dgvPileCaps == null) return;
            dgvPileCaps.Rows.Clear();
            foreach (var name in PileReactionChecker.GetSpringTypes(_sap))
                dgvPileCaps.Rows.Add(name, "", "");

            if (dgvPileCaps.Rows.Count == 0 && lblPileInfo != null)
                lblPileInfo.Text = "Model chưa khai báo loại point spring nào. Hãy gán point spring cho cọc trước.";
        }

        // Đồng bộ bảng SCT với các loại cọc thực sự phát hiện được trong model
        // (giữ nguyên giá trị SCT đã nhập, chỉ thêm dòng còn thiếu).
        private void SyncPileCapsGrid()
        {
            if (dgvPileCaps == null) return;

            var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (DataGridViewRow row in dgvPileCaps.Rows)
            {
                if (row.IsNewRow) continue;
                string nm = Convert.ToString(row.Cells[0].Value);
                if (!string.IsNullOrWhiteSpace(nm)) existing.Add(nm.Trim());
            }

            List<string> keys;
            try { keys = PileReactionChecker.GetPileTypeKeys(_sap); }
            catch { return; }

            foreach (var key in keys)
                if (!string.IsNullOrWhiteSpace(key) && !existing.Contains(key))
                {
                    dgvPileCaps.Rows.Add(key, "", "");
                    existing.Add(key);
                }
        }

        private Dictionary<string, PileSpringType> ReadPileCaps()
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
                    TensionCap = ParseCap(row.Cells[1].Value),
                    CompressionCap = ParseCap(row.Cells[2].Value)
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
            var caps = ReadPileCaps();
            var cases = new List<PileReactionCase>();
            AddPileCase(cases, "TỔ HỢP TẢI ĐỨNG", "TAI DUNG", cboPileVert.Text.Trim(), caps);
            AddPileCase(cases, "TỔ HỢP TẢI GIÓ", "TAI GIO", cboPileWind.Text.Trim(), caps);
            AddPileCase(cases, "TỔ HỢP TẢI ĐỘNG ĐẤT", "TAI DONG DAT", cboPileEq.Text.Trim(), caps);
            return cases;
        }

        private void AddPileCase(List<PileReactionCase> cases, string title, string sheet,
            string combo, Dictionary<string, PileSpringType> caps)
        {
            if (string.IsNullOrWhiteSpace(combo)) return;
            var rows = PileReactionChecker.Compute(_sap, combo, caps);
            cases.Add(new PileReactionCase { Title = title, SheetName = sheet, Combo = combo, Rows = rows });
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
