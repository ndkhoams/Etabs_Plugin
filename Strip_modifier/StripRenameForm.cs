using ETABSv1;
using Strip_Rename;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Etabs_Ultimate_Tools
{
    public class StripRenameForm : Form
    {
        private readonly cSapModel _sap;

        // Yeu cau chot: startIndex = 1; padding = 3 (001)
        private const int StartIndex = StripRenamerService.DefaultStartIndex;
        private const int PadWidth = StripRenamerService.DefaultPadWidth;

        private TextBox txtPrefix;
        private ComboBox cboSort;
        private CheckedListBox clbStrips;
        private Button btnReload, btnSelectAll, btnClear, btnPreview, btnApply, btnClose;
        private DataGridView dgv;
        private Label lblInfo;

        private List<StripInfo> _preview = new List<StripInfo>();

        public StripRenameForm(cSapModel sap)
        {
            _sap = sap;
            InitializeComponent();
            LoadStrips();
        }

        private void InitializeComponent()
        {
            Text = "ETABS Strip Rename  \u00A92026v1 by KhoaND13";
            Width = 940;
            Height = 640;
            MinimumSize = new Size(840, 560);
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Arial", 9F);

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 4,
                Padding = new Padding(12)
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 340));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 96));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            Controls.Add(root);

            var title = new Label
            {
                Text = "DOI TEN DESIGN STRIP",
                Dock = DockStyle.Fill,
                Font = new Font("Arial", 14F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter
            };
            root.Controls.Add(title, 0, 0);
            root.SetColumnSpan(title, 2);

            // ----- Options -----
            var optBox = new GroupBox { Dock = DockStyle.Fill, Text = "Thiet lap", Padding = new Padding(10, 6, 10, 6) };
            root.Controls.Add(optBox, 0, 1);
            root.SetColumnSpan(optBox, 2);

            var opt = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = true };
            optBox.Controls.Add(opt);

            opt.Controls.Add(MakeLabel("Prefix:", 56));
            txtPrefix = new TextBox { Text = "CS", Width = 120, Margin = new Padding(0, 7, 24, 0) };
            opt.Controls.Add(txtPrefix);

            opt.Controls.Add(MakeLabel("Kieu sap xep:", 90));
            cboSort = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 300, Margin = new Padding(0, 6, 24, 0) };
            cboSort.Items.AddRange(new object[]
            {
                "Trai->Phai, Tren->Duoi (LeftToRight_TopToBottom)",
                "Trai->Phai, Duoi->Tren (LeftToRight_BottomToTop)",
                "Phai->Trai, Tren->Duoi (RightToLeft_TopToBottom)",
                "Phai->Trai, Duoi->Tren (RightToLeft_BottomToTop)"
            });
            cboSort.SelectedIndex = 0;
            opt.Controls.Add(cboSort);

            opt.Controls.Add(new Label
            {
                Text = $"Index bat dau = {StartIndex}, so chu so = {PadWidth} (vi du: 001)",
                AutoSize = false, Width = 280, Height = 26,
                TextAlign = ContentAlignment.MiddleLeft, ForeColor = Color.DimGray,
                Margin = new Padding(0, 6, 0, 0)
            });

            // ----- Left: strip selection -----
            var stripsBox = new GroupBox { Dock = DockStyle.Fill, Text = "Chon strip can doi ten", Padding = new Padding(8) };
            root.Controls.Add(stripsBox, 0, 2);

            var stripsLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2 };
            stripsLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            stripsLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            stripsBox.Controls.Add(stripsLayout);

            var selBar = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = false };
            stripsLayout.Controls.Add(selBar, 0, 0);

            btnReload = MakeButton("Tai lai", 84); btnReload.Click += (s, e) => LoadStrips(); selBar.Controls.Add(btnReload);
            btnSelectAll = MakeButton("Chon tat ca", 96); btnSelectAll.Click += (s, e) => SetAllChecked(true); selBar.Controls.Add(btnSelectAll);
            btnClear = MakeButton("Bo chon", 84); btnClear.Click += (s, e) => SetAllChecked(false); selBar.Controls.Add(btnClear);

            clbStrips = new CheckedListBox { Dock = DockStyle.Fill, CheckOnClick = true, IntegralHeight = false };
            stripsLayout.Controls.Add(clbStrips, 0, 1);

            // ----- Right: preview grid -----
            var previewBox = new GroupBox { Dock = DockStyle.Fill, Text = "Xem truoc (cu -> moi)", Padding = new Padding(8) };
            root.Controls.Add(previewBox, 1, 2);

            dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                AllowUserToAddRows = false,
                ReadOnly = true,
                RowHeadersVisible = false,
                BackgroundColor = SystemColors.ControlLightLight,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                BorderStyle = BorderStyle.FixedSingle,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter },
                ColumnHeadersDefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            };
            previewBox.Controls.Add(dgv);
            AddColumns();

            // ----- Bottom bar -----
            var btnBar = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = false };
            root.Controls.Add(btnBar, 0, 3);
            root.SetColumnSpan(btnBar, 2);

            btnPreview = MakeButton("Xem truoc", 120); btnPreview.Click += (s, e) => RunPreview(); btnBar.Controls.Add(btnPreview);
            btnApply = MakeButton("Ap dung doi ten", 140); btnApply.Click += (s, e) => RunApply(); btnBar.Controls.Add(btnApply);
            btnClose = MakeButton("Dong", 90); btnClose.Click += (s, e) => Close(); btnBar.Controls.Add(btnClose);

            lblInfo = new Label
            {
                AutoSize = false, Width = 420, Height = 30,
                TextAlign = ContentAlignment.MiddleLeft, ForeColor = Color.DimGray,
                Margin = new Padding(16, 8, 0, 0)
            };
            btnBar.Controls.Add(lblInfo);
        }

        private void AddColumns()
        {
            dgv.Columns.Clear();
            dgv.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "OldName", HeaderText = "Ten cu", Width = 150, SortMode = DataGridViewColumnSortMode.NotSortable });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "NewName", HeaderText = "Ten moi", Width = 150, SortMode = DataGridViewColumnSortMode.NotSortable });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "X", HeaderText = "X (mid)", Width = 100, SortMode = DataGridViewColumnSortMode.NotSortable, DefaultCellStyle = { Format = "N3" } });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Y", HeaderText = "Y (mid)", Width = 100, SortMode = DataGridViewColumnSortMode.NotSortable, DefaultCellStyle = { Format = "N3" } });
        }

        private void LoadStrips()
        {
            try
            {
                var names = StripRenamerService.GetStripNames(_sap);
                clbStrips.Items.Clear();
                foreach (var n in names) clbStrips.Items.Add(n, true);

                _preview = new List<StripInfo>();
                dgv.DataSource = null;
                SetInfo($"Da tai {names.Count} strip.");
            }
            catch (Exception ex) { ShowError(ex); }
        }

        private void SetAllChecked(bool value)
        {
            for (int i = 0; i < clbStrips.Items.Count; i++) clbStrips.SetItemChecked(i, value);
        }

        private List<string> GetSelectedStrips()
        {
            return clbStrips.CheckedItems.Cast<object>().Select(o => o.ToString()).ToList();
        }

        private void RunPreview()
        {
            try
            {
                var selected = GetSelectedStrips();
                if (selected.Count == 0) { Warn("Hay chon it nhat 1 strip."); return; }

                string prefix = (txtPrefix.Text ?? string.Empty).Trim();
                var mode = (SortMode)cboSort.SelectedIndex;

                _preview = StripRenamerService.PreviewRename(_sap, prefix, StartIndex, PadWidth, mode, selected);

                dgv.DataSource = null;
                dgv.DataSource = _preview;
                SetInfo($"Xem truoc: {_preview.Count} strip se duoc doi ten.");
            }
            catch (Exception ex) { ShowError(ex); }
        }

        private void RunApply()
        {
            try
            {
                if (_preview == null || _preview.Count == 0) { Warn("Hay bam 'Xem truoc' truoc khi ap dung."); return; }

                var confirm = MessageBox.Show(
                    $"Se doi ten {_preview.Count} strip. Tiep tuc?",
                    "Xac nhan", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (confirm != DialogResult.Yes) return;

                StripRenamerService.ApplyRename(_sap, _preview);
                Info($"Da doi ten {_preview.Count} strip thanh cong.");
                LoadStrips();
            }
            catch (Exception ex) { ShowError(ex); }
        }

        // ---------- helpers ----------

        private void SetInfo(string text) => lblInfo.Text = text;

        private static Label MakeLabel(string text, int width) => new Label
        {
            Text = text, AutoSize = false, Width = width, Height = 26,
            TextAlign = ContentAlignment.MiddleLeft, Margin = new Padding(0, 6, 8, 0)
        };

        private static Button MakeButton(string text, int width) => new Button
        {
            Text = text, Width = width, Height = 28, Margin = new Padding(0, 8, 8, 0)
        };

        private static void Warn(string message) =>
            MessageBox.Show(message, "Strip Rename", MessageBoxButtons.OK, MessageBoxIcon.Warning);

        private static void Info(string message) =>
            MessageBox.Show(message, "Strip Rename", MessageBoxButtons.OK, MessageBoxIcon.Information);

        private static void ShowError(Exception ex) =>
            MessageBox.Show(ex.Message, "Strip Rename - Loi", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
}
