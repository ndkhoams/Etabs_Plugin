using ETABSv1;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace Etabs_Ultimate_Tools
{
    public partial class ModelCheckForm : Form
    {
        private TextBox txtStripGroup, txtStripPrefix, txtStripStart, txtStripPad, txtStripTol;
        private ComboBox cboStripTable, cboStripNameCol, cboStripXCol, cboStripYCol, cboStripSort;
        private Button btnStripRead, btnStripPreview, btnStripApply;
        private DataGridView dgvStrip;
        private Label lblStripInfo;
        private readonly StripRenamer _stripRenamer = new StripRenamer();
        private List<StripRenamer.StripItem> _stripPlan = new List<StripRenamer.StripItem>();

        private void BuildStripRenameTab(TabPage tab)
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 7, Padding = new Padding(12)
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 76));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 76));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            tab.Controls.Add(root);

            root.Controls.Add(MakeTitle("ĐỔI TÊN UNIQUE NAME CỦA DESIGN STRIP"), 0, 0);
            root.Controls.Add(MakeSubtitle("Đổi tên strip trong 1 Group theo prefix + số thứ tự, sắp xếp theo vị trí"), 0, 1);
            root.Controls.Add(MakeNote(
                "ETABS API không lấy được strip đang chọn → hãy chọn các strip trong ETABS rồi gán vào 1 Group " +
                "(Assign › Assign to Group), nhập tên Group vào đây. Nên chạy thử trên FILE COPY vì đổi tên strip " +
                "qua database có thể sinh strip trùng ở một số phiên bản ETABS."), 0, 2);

            // ----- Hàng tuỳ chọn -----
            var opt = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true, Margin = new Padding(0)
            };
            opt.Controls.Add(MakeFieldLabel("Group:", 48));
            txtStripGroup = MakeTextBox("", 130); opt.Controls.Add(txtStripGroup);
            opt.Controls.Add(MakeFieldLabel("Prefix:", 48));
            txtStripPrefix = MakeTextBox("S", 90); opt.Controls.Add(txtStripPrefix);
            opt.Controls.Add(MakeFieldLabel("Bắt đầu:", 60));
            txtStripStart = MakeTextBox("1", 50); opt.Controls.Add(txtStripStart);
            opt.Controls.Add(MakeFieldLabel("Số chữ số:", 72));
            txtStripPad = MakeTextBox("3", 40); opt.Controls.Add(txtStripPad);
            opt.Controls.Add(MakeFieldLabel("Dung sai gộp hàng (m):", 152));
            txtStripTol = MakeTextBox("0.5", 50); opt.Controls.Add(txtStripTol);
            opt.Controls.Add(MakeFieldLabel("Sắp xếp:", 60));
            cboStripSort = MakeCombo(240);
            cboStripSort.Items.AddRange(new object[]
            {
                "Hàng ngang: Trái→Phải, Trên→Dưới",
                "Hàng ngang ngược: Phải→Trái, Dưới→Trên",
                "Cột dọc: Trên→Dưới, Trái→Phải",
                "Cột dọc ngược: Dưới→Trên, Phải→Trái"
            });
            cboStripSort.SelectedIndex = 0;
            opt.Controls.Add(cboStripSort);
            root.Controls.Add(opt, 0, 3);

            // ----- Hàng ánh xạ cột + bảng + nút -----
            var mapPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true, Margin = new Padding(0)
            };
            mapPanel.Controls.Add(MakeFieldLabel("Bảng strip:", 72));
            cboStripTable = new ComboBox { DropDownStyle = ComboBoxStyle.DropDown, Width = 250, Margin = new Padding(0, 6, 18, 0) };
            mapPanel.Controls.Add(cboStripTable);
            mapPanel.Controls.Add(MakeFieldLabel("Cột tên:", 56));
            cboStripNameCol = MakeCombo(150); mapPanel.Controls.Add(cboStripNameCol);
            mapPanel.Controls.Add(MakeFieldLabel("Cột X:", 46));
            cboStripXCol = MakeCombo(120); mapPanel.Controls.Add(cboStripXCol);
            mapPanel.Controls.Add(MakeFieldLabel("Cột Y:", 46));
            cboStripYCol = MakeCombo(120); mapPanel.Controls.Add(cboStripYCol);

            btnStripRead = MakeButton("Đọc bảng"); btnStripRead.Click += (s, e) => StripReadTable(); mapPanel.Controls.Add(btnStripRead);
            btnStripPreview = MakeButton("Xem trước"); btnStripPreview.Click += (s, e) => StripPreview(); mapPanel.Controls.Add(btnStripPreview);
            btnStripApply = MakeButton("Áp dụng đổi tên"); btnStripApply.Enabled = false; btnStripApply.Click += (s, e) => StripApply(); mapPanel.Controls.Add(btnStripApply);
            root.Controls.Add(mapPanel, 0, 4);

            // ----- Lưới xem trước -----
            dgvStrip = CreateGrid();
            AddColumn(dgvStrip, "Stt", "STT", 50);
            AddColumn(dgvStrip, "OldName", "Tên hiện tại", 210);
            AddColumn(dgvStrip, "NewName", "Tên mới", 210);
            AddColumn(dgvStrip, "X", "X", 110, "0.000");
            AddColumn(dgvStrip, "Y", "Y", 110, "0.000");
            AddColumn(dgvStrip, "RecordCount", "Số dòng", 80);
            root.Controls.Add(dgvStrip, 0, 5);

            lblStripInfo = new Label
            {
                Dock = DockStyle.Fill, ForeColor = Color.DimGray,
                TextAlign = ContentAlignment.MiddleLeft, AutoSize = false
            };
            root.Controls.Add(lblStripInfo, 0, 6);
        }

        private void StripReadTable()
        {
            try
            {
                if (cboStripTable.Items.Count == 0) StripLoadTableList();
                string tableKey = (cboStripTable.Text ?? "").Trim();
                if (string.IsNullOrWhiteSpace(tableKey)) { Warn("Chọn hoặc nhập tên bảng design strip.", "Đổi tên Strip"); return; }
                string group = (txtStripGroup.Text ?? "").Trim();
                if (string.IsNullOrWhiteSpace(group)) { Warn("Nhập tên Group chứa các strip cần đổi tên.", "Đổi tên Strip"); return; }

                _stripRenamer.ReadGroup(_sap, tableKey, group);
                StripPopulateColumnCombos();

                if (_stripRenamer.NumberRecords == 0)
                {
                    dgvStrip.DataSource = null;
                    btnStripApply.Enabled = false;
                    lblStripInfo.Text = "Đọc được 0 dòng cho group '" + group + "'. Kiểm tra lại tên bảng / tên group.";
                    return;
                }
                StripPreview();
            }
            catch (Exception ex) { Warn(ex.Message, "Đổi tên Strip"); }
        }

        private void StripLoadTableList()
        {
            string[] guesses =
            {
                "Object Geometry - Design Strips",
                "Object Connectivity - Design Strips",
                "Design Strip Object Geometry",
                "Slab Design - Strip Definitions"
            };
            foreach (var g in guesses)
                if (!cboStripTable.Items.Contains(g)) cboStripTable.Items.Add(g);
            try
            {
                var found = StripRenamer.FindStripTables(_sap);
                foreach (var kv in found)
                    if (!cboStripTable.Items.Contains(kv.Key)) cboStripTable.Items.Add(kv.Key);
            }
            catch { /* GetAllTables có thể không khả dụng – dùng danh sách phỏng đoán */ }
            if (string.IsNullOrWhiteSpace(cboStripTable.Text) && cboStripTable.Items.Count > 0)
                cboStripTable.SelectedIndex = 0;
        }

        private void StripPopulateColumnCombos()
        {
            cboStripNameCol.Items.Clear();
            cboStripXCol.Items.Clear();
            cboStripYCol.Items.Clear();
            cboStripXCol.Items.Add("(không)");
            cboStripYCol.Items.Add("(không)");
            foreach (var f in _stripRenamer.FieldKeys)
            {
                cboStripNameCol.Items.Add(f);
                cboStripXCol.Items.Add(f);
                cboStripYCol.Items.Add(f);
            }
            cboStripNameCol.SelectedIndex = _stripRenamer.NameCol >= 0 && _stripRenamer.NameCol < cboStripNameCol.Items.Count
                ? _stripRenamer.NameCol : (cboStripNameCol.Items.Count > 0 ? 0 : -1);
            cboStripXCol.SelectedIndex = _stripRenamer.XCol >= 0 ? _stripRenamer.XCol + 1 : 0;
            cboStripYCol.SelectedIndex = _stripRenamer.YCol >= 0 ? _stripRenamer.YCol + 1 : 0;
        }

        private void StripPreview()
        {
            if (_stripRenamer.NumberRecords == 0) { Warn("Bấm 'Đọc bảng' trước.", "Đổi tên Strip"); return; }
            int nameCol = cboStripNameCol.SelectedIndex;
            if (nameCol < 0) { Warn("Chọn cột chứa tên strip.", "Đổi tên Strip"); return; }
            int xCol = cboStripXCol.SelectedIndex - 1;
            int yCol = cboStripYCol.SelectedIndex - 1;
            string prefix = txtStripPrefix.Text ?? "";
            int start = ParseIntOr(txtStripStart.Text, 1);
            int pad = ParseIntOr(txtStripPad.Text, 3);
            double tol = ParseDoubleOr(txtStripTol.Text, 0.5);
            var mode = (StripRenamer.SortMode)Math.Max(0, cboStripSort.SelectedIndex);
            try
            {
                _stripPlan = _stripRenamer.BuildPlan(nameCol, xCol, yCol, prefix, start, pad, mode, tol);
                var rows = new List<StripPreviewRow>();
                int stt = 1;
                foreach (var it in _stripPlan)
                    rows.Add(new StripPreviewRow
                    {
                        Stt = stt++, OldName = it.OldName, NewName = it.NewName,
                        X = it.X, Y = it.Y, RecordCount = it.RecordCount
                    });
                dgvStrip.DataSource = null;
                dgvStrip.DataSource = rows;
                bool hasCoord = xCol >= 0 && yCol >= 0;
                lblStripInfo.Text = "Tìm thấy " + _stripPlan.Count + " strip trong group '" + _stripRenamer.GroupName + "'."
                    + (hasCoord ? "" : "  —  CHƯA chọn cột toạ độ X/Y nên sắp xếp theo tên (không theo vị trí).");
                btnStripApply.Enabled = _stripPlan.Count > 0;
            }
            catch (Exception ex) { Warn(ex.Message, "Đổi tên Strip"); }
        }

        private void StripApply()
        {
            if (_stripPlan == null || _stripPlan.Count == 0) { Warn("Chưa có dữ liệu xem trước.", "Đổi tên Strip"); return; }
            int nameCol = cboStripNameCol.SelectedIndex;
            if (nameCol < 0) { Warn("Chọn cột chứa tên strip.", "Đổi tên Strip"); return; }
            var confirm = MessageBox.Show(
                "Sẽ đổi tên " + _stripPlan.Count + " strip trong group '" + _stripRenamer.GroupName + "'.\n\n" +
                "CẢNH BÁO: đổi tên strip qua database có thể sinh strip trùng ở một số phiên bản ETABS, " +
                "và thao tác này sẽ MỞ KHÓA model (xóa kết quả phân tích nếu có). " +
                "Hãy chắc chắn bạn đang chạy trên FILE COPY.\n\nTiếp tục?",
                "Xác nhận đổi tên Strip", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (confirm != DialogResult.Yes) return;
            try
            {
                string log = _stripRenamer.Apply(_sap, _stripPlan, nameCol);
                Info("Đã áp dụng đổi tên.\n\n" + log, "Đổi tên Strip");
                btnStripApply.Enabled = false;
                lblStripInfo.Text = "Đã áp dụng. Bấm 'Đọc bảng' để tải lại danh sách theo tên mới.";
            }
            catch (Exception ex) { Warn(ex.Message, "Đổi tên Strip"); }
        }

        private static int ParseIntOr(string s, int def)
        {
            int v;
            return int.TryParse((s ?? "").Trim(), out v) ? v : def;
        }

        private static double ParseDoubleOr(string s, double def)
        {
            double v;
            string t = (s ?? "").Trim();
            if (double.TryParse(t, NumberStyles.Any, CultureInfo.InvariantCulture, out v)) return v;
            if (double.TryParse(t.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out v)) return v;
            return def;
        }
    }
}
