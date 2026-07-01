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
        // Bề rộng "chuẩn" của bảng SCT khi hiện đủ 9 cột — luôn giữ cố định để bảng lấp kín box.
        private const int CapHFullW = CapHNameW + 9 * CapHValW + 4;

        // Mỗi trường hợp tải cho chọn NHIỀU tổ hợp; plug-in xét tất cả và lấy combo nguy hiểm nhất.
        private CheckedListBox clbPileHVert, clbPileHWind, clbPileHEq;
        private DataGridView dgvPileHCaps, dgvPileHPreview;
        private Button btnPileHPreview, btnPileHExport;
        private Label lblPileHInfo;
        private CheckBox chkPileHConsiderH, chkPileHConsiderTension, chkPileHConsiderCompression;
        private TableLayoutPanel _pileHCapsPanel;
        private Control _pileHCapsHeader;
        private List<PileReactionCase> _pileHCases = new List<PileReactionCase>();

        // Lưu chỉ số click gần nhất của từng danh sách để hỗ trợ giữ SHIFT chọn dải.
        private readonly Dictionary<CheckedListBox, int> _pileHLastClick = new Dictionary<CheckedListBox, int>();

        // Có xét tải ngang hay không (mặc định có).
        private bool ConsiderPileH
        {
            get { return chkPileHConsiderH == null || chkPileHConsiderH.Checked; }
        }

        // Có xét SCT chịu kéo hay không (mặc định có).
        private bool ConsiderPileTension
        {
            get { return chkPileHConsiderTension == null || chkPileHConsiderTension.Checked; }
        }

        // Có xét SCT chịu nén hay không (mặc định có).
        private bool ConsiderPileCompression
        {
            get { return chkPileHConsiderCompression == null || chkPileHConsiderCompression.Checked; }
        }

        // ---------- Helper UI dùng chung cho tab Pile Reactions (gộp từ PileShared) ----------

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

        private static CheckedListBox MakePileCheckList() => new CheckedListBox
        {
            Dock = DockStyle.Fill, CheckOnClick = true, IntegralHeight = false,
            BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(0, 0, 6, 0)
        };

        // Nút "Bỏ chọn" cho 1 danh sách tổ hợp.
        private Button MakePileClearButton(CheckedListBox clb)
        {
            var b = new Button
            {
                Text = "Bỏ chọn", Dock = DockStyle.Fill, Height = 24,
                Margin = new Padding(0, 2, 6, 0)
            };
            b.Click += (s, e) => SetPileListChecked(clb, false);
            return b;
        }

        private static void SetPileListChecked(CheckedListBox clb, bool state)
        {
            if (clb == null) return;
            for (int i = 0; i < clb.Items.Count; i++) clb.SetItemChecked(i, state);
        }

        // Giữ SHIFT + click để tích/bỏ tích một dải tổ hợp liên tiếp (giống tab Column Force).
        private void EnablePileHShiftSelect(CheckedListBox clb)
        {
            if (clb == null) return;
            _pileHLastClick[clb] = -1;
            clb.MouseDown += (s, e) =>
            {
                int index = clb.IndexFromPoint(e.Location);
                if (index < 0) return;
                int last;
                if (!_pileHLastClick.TryGetValue(clb, out last)) last = -1;
                if ((Control.ModifierKeys & Keys.Shift) == Keys.Shift && last >= 0 && last != index)
                {
                    bool target = !clb.GetItemChecked(index);
                    int start = Math.Min(last, index);
                    int end = Math.Max(last, index);
                    for (int i = start; i <= end; i++)
                    {
                        if (i == index) continue;
                        clb.SetItemChecked(i, target);
                    }
                }
                _pileHLastClick[clb] = index;
            };
        }

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
                Margin = new Padding(0, 8, 0, 0),
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter },
                ColumnHeadersDefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            };
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

        private static double ParseCap(object value)
        {
            if (value == null) return 0.0;
            string s = value.ToString().Trim().Replace(",", ".");
            double d;
            return double.TryParse(s, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out d) ? d : 0.0;
        }

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

            root.Controls.Add(MakeTitle("KIỂM TRA KHẢ NĂNG CHỊU TẢI CỦA CỌC"), 0, 0);
            root.Controls.Add(MakeSubtitle("(Phản lực đứng so với SCT kéo/nén; hợp lực ngang H=√(FX²+FY²) so với SCT ngang, đơn vị kN)"), 0, 1);

            var main = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Margin = new Padding(0, 6, 0, 0)
            };
            main.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 510));
            main.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.Controls.Add(main, 0, 2);

            var left = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 6, Margin = new Padding(0, 0, 10, 0)
            };
            left.RowStyles.Add(new RowStyle(SizeType.Absolute, 178)); // 3 danh sách chọn tổ hợp + nút Bỏ chọn
            left.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));  // hàng tùy chọn (3 checkbox)
            left.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));  // nhãn SCT
            left.RowStyles.Add(new RowStyle(SizeType.Percent, 100));  // bảng SCT
            left.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));  // nút
            left.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));  // thông tin
            main.Controls.Add(left, 0, 0);

            // Mỗi cột là 1 danh sách tích chọn nhiều tổ hợp cho 1 trường hợp tải + nút Bỏ chọn.
            var comboPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 3
            };
            comboPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));
            comboPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));
            comboPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.34f));
            comboPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 20));
            comboPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            comboPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            left.Controls.Add(comboPanel, 0, 0);

            comboPanel.Controls.Add(MakePileLabel("Tổ hợp tải đứng:"), 0, 0);
            comboPanel.Controls.Add(MakePileLabel("Tổ hợp tải gió:"), 1, 0);
            comboPanel.Controls.Add(MakePileLabel("Tổ hợp tải động đất:"), 2, 0);
            clbPileHVert = MakePileCheckList(); EnablePileHShiftSelect(clbPileHVert); comboPanel.Controls.Add(clbPileHVert, 0, 1);
            clbPileHWind = MakePileCheckList(); EnablePileHShiftSelect(clbPileHWind); comboPanel.Controls.Add(clbPileHWind, 1, 1);
            clbPileHEq = MakePileCheckList(); EnablePileHShiftSelect(clbPileHEq); comboPanel.Controls.Add(clbPileHEq, 2, 1);
            comboPanel.Controls.Add(MakePileClearButton(clbPileHVert), 0, 2);
            comboPanel.Controls.Add(MakePileClearButton(clbPileHWind), 1, 2);
            comboPanel.Controls.Add(MakePileClearButton(clbPileHEq), 2, 2);

            // Hàng tùy chọn: 3 checkbox độc lập (kiểm tra cọc nén / kéo / tải ngang).
            var optRow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = true,
                Margin = new Padding(0)
            };
            left.Controls.Add(optRow, 0, 1);

            chkPileHConsiderCompression = new CheckBox
            {
                Text = "Kiểm tra cọc chịu nén",
                Checked = true,
                AutoSize = true,
                Margin = new Padding(0, 4, 16, 0)
            };
            chkPileHConsiderCompression.CheckedChanged += (s, e) => ApplyPileHColumnVisibility();
            optRow.Controls.Add(chkPileHConsiderCompression);

            chkPileHConsiderTension = new CheckBox
            {
                Text = "Kiểm tra cọc chịu kéo",
                Checked = true,
                AutoSize = true,
                Margin = new Padding(0, 4, 16, 0)
            };
            chkPileHConsiderTension.CheckedChanged += (s, e) => ApplyPileHColumnVisibility();
            optRow.Controls.Add(chkPileHConsiderTension);

            chkPileHConsiderH = new CheckBox
            {
                Text = "Kiểm tra cọc chịu tải ngang",
                Checked = true,
                AutoSize = true,
                Margin = new Padding(0, 4, 0, 0)
            };
            chkPileHConsiderH.CheckedChanged += (s, e) => ApplyPileHColumnVisibility();
            optRow.Controls.Add(chkPileHConsiderH);

            left.Controls.Add(new Label
            {
                Text = "SCT theo loại cọc & từng tổ hợp (kN):",
                Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft
            }, 0, 2);

            // Panel bảng SCT: dùng Anchor Top|Left, GIỮ bề rộng cố định = CapHFullW (lấp kín box).
            _pileHCapsPanel = new TableLayoutPanel
            {
                Anchor = AnchorStyles.Top | AnchorStyles.Left,
                ColumnCount = 1, RowCount = 2,
                Width = CapHFullW, Margin = new Padding(0, 8, 0, 0)
            };
            _pileHCapsPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
            _pileHCapsPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            left.Controls.Add(_pileHCapsPanel, 0, 3);

            dgvPileHCaps = CreateEditableGrid();
            dgvPileHCaps.ColumnHeadersVisible = false;
            dgvPileHCaps.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            dgvPileHCaps.AllowUserToResizeColumns = false;
            dgvPileHCaps.AllowUserToResizeRows = false;
            dgvPileHCaps.ScrollBars = ScrollBars.None;
            dgvPileHCaps.Margin = new Padding(0);
            _pileHCapsPanel.Controls.Add(dgvPileHCaps, 0, 1);
            AddPileHCapsColumns();

            var btnRow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = false,
                Margin = new Padding(0, 6, 0, 0)
            };
            left.Controls.Add(btnRow, 0, 4);

            btnPileHPreview = new Button { Text = "Kiểm tra", Width = 130, Height = 38, Margin = new Padding(0, 0, 12, 0) };
            btnPileHPreview.Click += (s, e) => PreviewPileHReactions();
            btnRow.Controls.Add(btnPileHPreview);

            btnPileHExport = new Button { Text = "Xuất Excel", Width = 150, Height = 38, Enabled = false, Margin = new Padding(0, 0, 0, 0) };
            btnPileHExport.Click += (s, e) => ExportPileHReactions();
            btnRow.Controls.Add(btnPileHExport);

            lblPileHInfo = new Label
            {
                Dock = DockStyle.Fill, ForeColor = Color.DimGray, TextAlign = ContentAlignment.MiddleLeft
            };
            left.Controls.Add(lblPileHInfo, 0, 5);

            var right = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2
            };
            right.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
            right.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            main.Controls.Add(right, 1, 0);

            dgvPileHPreview = CreateGrid();
            right.Controls.Add(dgvPileHPreview, 0, 1);
            AddPileHGridColumns();

            right.Controls.Add(new Label
            {
                Text = "Preview:",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            }, 0, 0);

            lblPileHInfo.Text = "Điền SCT cọc (kéo/nén/ngang) vào bảng.";

            // Dựng header bảng SCT + đồng bộ ẩn/hiện cột theo trạng thái checkbox.
            ApplyPileHColumnVisibility();
        }

        // Header gộp 2 dòng: dòng trên = trường hợp tải, dòng dưới = Kéo/Nén/Ngang.
        // Mỗi nhóm hiện: Kéo (nếu considerTension), Nén (nếu considerCompression), Ngang (nếu considerH).
        // valWidths: bề rộng từng cột giá trị (đã chia để lấp kín bề rộng bảng).
        private Control BuildCapsHeaderH(bool considerTension, bool considerCompression, bool considerH, int[] valWidths)
        {
            int perGroup = (considerTension ? 1 : 0) + (considerCompression ? 1 : 0) + (considerH ? 1 : 0);
            int valCols = perGroup * 3;

            var h = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 1 + valCols, RowCount = 2, Margin = new Padding(0)
            };
            h.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, CapHNameW));
            for (int i = 0; i < valCols; i++) h.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, valWidths[i]));
            h.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
            h.RowStyles.Add(new RowStyle(SizeType.Percent, 50));

            var name = MakeHeadCell("Loại cọc");
            h.Controls.Add(name, 0, 0);
            h.SetRowSpan(name, 2);

            if (perGroup > 0)
            {
                string[] groups = { "Tải đứng", "Tải gió", "Tải động đất" };
                int col = 1;
                for (int g = 0; g < 3; g++)
                {
                    var gc = MakeHeadCell(groups[g]);
                    h.Controls.Add(gc, col, 0);
                    h.SetColumnSpan(gc, perGroup);

                    int sub = col;
                    if (considerTension) { h.Controls.Add(MakeHeadCell("Kéo"), sub, 1); sub++; }
                    if (considerCompression) { h.Controls.Add(MakeHeadCell("Nén"), sub, 1); sub++; }
                    if (considerH) { h.Controls.Add(MakeHeadCell("Ngang"), sub, 1); sub++; }

                    col += perGroup;
                }
            }
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
            AddColumn(dgvPileHPreview, "Combo", "Tổ hợp", 120, null, true);
            AddColumn(dgvPileHPreview, "Reaction", "FZ (kN)", 90, "0.#", true);
            AddColumn(dgvPileHPreview, "TensionCap", "SCT kéo", 70, "0.#", true);
            AddColumn(dgvPileHPreview, "CompressionCap", "SCT nén", 70, "0.#", true);
            AddColumn(dgvPileHPreview, "Result", "KL SCT\nđứng", 75, null, true);
            AddColumn(dgvPileHPreview, "Fx", "|FX| (kN)", 70, "0.#", true);
            AddColumn(dgvPileHPreview, "Fy", "|FY| (kN)", 70, "0.#", true);
            AddColumn(dgvPileHPreview, "Horizontal", "H (kN)", 70, "0.#", true);
            AddColumn(dgvPileHPreview, "HorizontalCap", "SCT ngang", 75, "0.#", true);
            AddColumn(dgvPileHPreview, "HResult", "KL SCT\nngang", 75, null, true);
        }

        // Ẩn/hiện cột liên quan SCT nén/kéo & tải ngang (bảng SCT, header, bảng preview) theo checkbox.
        private void ApplyPileHColumnVisibility()
        {
            bool considerH = ConsiderPileH;
            bool considerT = ConsiderPileTension;
            bool considerC = ConsiderPileCompression;
            bool anyVert = considerT || considerC;

            // 1) Bảng nhập SCT: ẩn/hiện cột Kéo (1,4,7), Nén (2,5,8) và Ngang (3,6,9).
            if (dgvPileHCaps != null && dgvPileHCaps.Columns.Count >= 10)
            {
                dgvPileHCaps.Columns[CapHTensVert].Visible = considerT;
                dgvPileHCaps.Columns[CapHTensWind].Visible = considerT;
                dgvPileHCaps.Columns[CapHTensEq].Visible = considerT;
                dgvPileHCaps.Columns[CapHCompVert].Visible = considerC;
                dgvPileHCaps.Columns[CapHCompWind].Visible = considerC;
                dgvPileHCaps.Columns[CapHCompEq].Visible = considerC;
                dgvPileHCaps.Columns[CapHHorizVert].Visible = considerH;
                dgvPileHCaps.Columns[CapHHorizWind].Visible = considerH;
                dgvPileHCaps.Columns[CapHHorizEq].Visible = considerH;
            }

            // 2) Dựng lại header bảng SCT cho khớp số cột hiển thị (giữ bề rộng = CapHFullW).
            if (_pileHCapsPanel != null)
            {
                _pileHCapsPanel.SuspendLayout();
                if (_pileHCapsHeader != null)
                {
                    _pileHCapsPanel.Controls.Remove(_pileHCapsHeader);
                    _pileHCapsHeader.Dispose();
                }

                int perGroup = (considerT ? 1 : 0) + (considerC ? 1 : 0) + (considerH ? 1 : 0);
                int valCols = perGroup * 3;

                int[] valWidths;
                if (valCols > 0)
                {
                    int avail = CapHFullW - CapHNameW - 4;
                    int baseW = avail / valCols;
                    int rem = avail - baseW * valCols;
                    valWidths = new int[valCols];
                    for (int i = 0; i < valCols; i++) valWidths[i] = baseW + (i < rem ? 1 : 0);
                }
                else
                {
                    valWidths = new int[0];
                }

                _pileHCapsHeader = BuildCapsHeaderH(considerT, considerC, considerH, valWidths);
                _pileHCapsPanel.Controls.Add(_pileHCapsHeader, 0, 0);

                // Đặt lại bề rộng các cột đang hiển thị trong lưới nhập cho khớp header.
                if (dgvPileHCaps != null && dgvPileHCaps.Columns.Count >= 10)
                {
                    dgvPileHCaps.Columns[0].Width = CapHNameW;
                    int vi = 0;
                    for (int i = 1; i < dgvPileHCaps.Columns.Count; i++)
                    {
                        if (!dgvPileHCaps.Columns[i].Visible) continue;
                        if (vi < valWidths.Length) dgvPileHCaps.Columns[i].Width = valWidths[vi];
                        vi++;
                    }
                }

                // LUÔN giữ bề rộng đầy đủ — KHÔNG co theo số cột.
                _pileHCapsPanel.Width = CapHFullW;
                _pileHCapsPanel.ResumeLayout(true);
                if (_pileHCapsPanel.Parent != null) _pileHCapsPanel.Parent.PerformLayout();
            }

            // 3) Bảng preview: ẩn/hiện FZ (4), SCT kéo (5), SCT nén (6), KL đứng (7) và các cột ngang (8..12).
            if (dgvPileHPreview != null)
            {
                if (4 < dgvPileHPreview.Columns.Count) dgvPileHPreview.Columns[4].Visible = anyVert;
                if (5 < dgvPileHPreview.Columns.Count) dgvPileHPreview.Columns[5].Visible = considerT;
                if (6 < dgvPileHPreview.Columns.Count) dgvPileHPreview.Columns[6].Visible = considerC;
                if (7 < dgvPileHPreview.Columns.Count) dgvPileHPreview.Columns[7].Visible = anyVert;

                int[] hCols = { 8, 9, 10, 11, 12 };
                foreach (int ci in hCols)
                    if (ci < dgvPileHPreview.Columns.Count)
                        dgvPileHPreview.Columns[ci].Visible = considerH;
            }

            // 4) Co khít chiều cao panel SCT theo số dòng hiện có.
            AdjustPileHCapsHeight();
        }

        // Co chiều cao panel SCT vừa khít nội dung (tránh vùng trắng thừa bên dưới).
        private void AdjustPileHCapsHeight()
        {
            if (_pileHCapsPanel == null || dgvPileHCaps == null) return;
            int gridH = 4;
            foreach (DataGridViewRow r in dgvPileHCaps.Rows)
                if (!r.IsNewRow) gridH += r.Height;
            if (gridH < 28) gridH = 28;
            const int maxH = 360;
            if (gridH > maxH)
            {
                gridH = maxH;
                dgvPileHCaps.ScrollBars = ScrollBars.Vertical;
            }
            else
            {
                dgvPileHCaps.ScrollBars = ScrollBars.None;
            }
            _pileHCapsPanel.Height = 46 + gridH + 2;
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

            AdjustPileHCapsHeight();
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

            AdjustPileHCapsHeight();
        }

        // Điền SCT tạm = Kz×0.01 vào các ô KÉO/NÉN còn trống; cột NGANG lấy tạm = 1/10 SCT nén.
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

            // Lực ngang lấy tạm = 1/10 SCT nén; chỉ điền vào ô NGANG còn trống.
            string horizVal = Math.Round(defaultCap / 10.0, 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
            int[] horizCols = { CapHHorizVert, CapHHorizWind, CapHHorizEq };
            foreach (int c in horizCols)
            {
                object cur = row.Cells[c].Value;
                if (cur == null || string.IsNullOrWhiteSpace(cur.ToString()))
                    row.Cells[c].Value = horizVal;
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
                    TensionCap = ConsiderPileTension ? ParseCap(row.Cells[tensCol].Value) : 0.0,
                    CompressionCap = ConsiderPileCompression ? ParseCap(row.Cells[compCol].Value) : 0.0,
                    HorizontalCap = ConsiderPileH ? ParseCap(row.Cells[horizCol].Value) : 0.0
                };
            }
            return dict;
        }

        private List<PileReactionCase> BuildPileHCases()
        {
            var cases = new List<PileReactionCase>();
            AddPileHCase(cases, GetCheckedCombos(clbPileHVert), "TẢI ĐỨNG", "TAI DUNG",
                ReadPileHCaps(CapHTensVert, CapHCompVert, CapHHorizVert));
            AddPileHCase(cases, GetCheckedCombos(clbPileHWind), "TẢI GIÓ", "TAI GIO",
                ReadPileHCaps(CapHTensWind, CapHCompWind, CapHHorizWind));
            AddPileHCase(cases, GetCheckedCombos(clbPileHEq), "ĐỘNG ĐẤT", "TAI DONG DAT",
                ReadPileHCaps(CapHTensEq, CapHCompEq, CapHHorizEq));
            return cases;
        }

        // Lấy danh sách tổ hợp được tích chọn trong 1 CheckedListBox.
        private static List<string> GetCheckedCombos(CheckedListBox clb)
        {
            var list = new List<string>();
            if (clb == null) return list;
            foreach (var item in clb.CheckedItems)
            {
                string s = Convert.ToString(item);
                if (!string.IsNullOrWhiteSpace(s)) list.Add(s.Trim());
            }
            return list;
        }

        private void AddPileHCase(List<PileReactionCase> cases, List<string> combos, string title,
            string sheet, Dictionary<string, PileSpringType> caps)
        {
            if (combos == null || combos.Count == 0) return;
            var c = PileReactionChecker.ComputeCaseHMulti(_sap, combos, title, sheet, caps,
                ConsiderPileTension, ConsiderPileCompression);
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
                Warn("Chưa chọn tổ hợp tải nào (cần tích ít nhất 1 tổ hợp trong 1 trong 3 trường hợp).", PileHTitle);
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
            ApplyPileHColumnVisibility();

            bool considerH = ConsiderPileH;
            bool anyVert = ConsiderPileTension || ConsiderPileCompression;
            int fail = preview.Count(p =>
                (anyVert && !string.IsNullOrEmpty(p.Result) && p.Result.IndexOf("Không", StringComparison.OrdinalIgnoreCase) >= 0)
                || (considerH && !string.IsNullOrEmpty(p.HResult) && p.HResult.IndexOf("Không", StringComparison.OrdinalIgnoreCase) >= 0));

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
                    PileReactionExporterH.Export(sfd.FileName, _pileHCases,
                        ConsiderPileTension, ConsiderPileCompression, ConsiderPileH);
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
