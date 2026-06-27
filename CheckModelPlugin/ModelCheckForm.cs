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

        private ComboBox cboAxialConcrete, cboAxialCombo;
        private Button btnAxialRun, btnAxialExport;
        private DataGridView dgvAxial;
        private Label lblAxialInfo;
        private List<AxialCheckRow> _axialRows = new List<AxialCheckRow>();

        private CheckedListBox clbColCombos;
        private Button btnColPreview, btnColExportText, btnColExportExcel;
        private DataGridView dgvColPreview;
        private Label lblColInfo;
        private List<ForceRow> _colRows = new List<ForceRow>();
        private int _lastColIndex = -1;

        // Property Modifiers (mỗi nhóm cấu kiện là 1 ModGroup tái sử dụng)
        private ModGroup _modBeam, _modCol, _modSlab, _modWall;
        private Button btnModApply, btnModRollback;
        private Label lblModInfo;

        private const double AxialAlphaCc = 1.0;
        private const double AxialGammaC = 1.2;
        private const double AxialColumnLimit = 0.65;
        private const double AxialWallLimit = 0.40;

        private const double WindDriftLimitDen = 500.0;

        public ModelCheckForm(cSapModel sap)
        {
            _sap = sap;
            InitializeComponent();
            LoadCombos();
        }

        private void InitializeComponent()
        {
            Text = "Etabs Ultimate Tools  ©2026v1 by KhoaND13";
            Width = 1480;
            Height = 780;
            MinimumSize = new Size(1360, 700);
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Arial", 9F);

            var tabs = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = new Font("Arial", 9F),
                SizeMode = TabSizeMode.Fixed,
                ItemSize = new Size(150, 38),
                DrawMode = TabDrawMode.OwnerDrawFixed,
                Padding = new Point(10, 4)
            };
            tabs.DrawItem += Tabs_DrawItem;
            Controls.Add(tabs);

            var tabModifier = new TabPage("Property Modifiers");
            var tabWind = new TabPage("Displacements");
            var tabWindDrift = new TabPage("Wind Drifts");
            var tabSeis = new TabPage("Seismic Drifts");
            var tabPDelta = new TabPage("P-Delta");
            var tabAxial = new TabPage("Axial Force");
            var tabColExport = new TabPage("Column Force Exporter");
            var tabPileH = new TabPage("Pile Reactions");

            tabs.TabPages.Add(tabModifier);
            tabs.TabPages.Add(tabWind);
            tabs.TabPages.Add(tabWindDrift);
            tabs.TabPages.Add(tabSeis);
            tabs.TabPages.Add(tabPDelta);
            tabs.TabPages.Add(tabAxial);
            tabs.TabPages.Add(tabColExport);
            tabs.TabPages.Add(tabPileH);

            BuildModifierTab(tabModifier);
            BuildWindTab(tabWind);
            BuildWindDriftTab(tabWindDrift);
            BuildSeismicDriftTab(tabSeis);
            BuildPDeltaTab(tabPDelta);
            BuildAxialTab(tabAxial);
            BuildColumnExportTab(tabColExport);
            BuildPileHTab(tabPileH);
        }

        // ---------- Hộp thoại dùng chung ----------

        private static void Warn(string message, string title) =>
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);

        private static void Info(string message, string title) =>
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Information);

        private static bool RequireCombo(ComboBox cbo, string title, string message, out string combo)
        {
            combo = cbo.Text.Trim();
            if (string.IsNullOrWhiteSpace(combo))
            {
                Warn(message, title);
                return false;
            }
            return true;
        }

        // ---------- Vẽ tab nổi bật (owner-draw) ----------

        private void Tabs_DrawItem(object sender, DrawItemEventArgs e)
        {
            var tc = (TabControl)sender;
            Rectangle tabRect = tc.GetTabRect(e.Index);
            bool selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;

            Color back = selected ? Color.FromArgb(37, 99, 235) : Color.FromArgb(226, 232, 240);
            Color fore = selected ? Color.White : Color.FromArgb(45, 55, 72);

            using (var b = new SolidBrush(back))
                e.Graphics.FillRectangle(b, tabRect);

            using (var tabFont = new Font("Arial", 9F, FontStyle.Bold))
                TextRenderer.DrawText(e.Graphics, tc.TabPages[e.Index].Text, tabFont, tabRect, fore,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);

            using (var pen = new Pen(selected ? Color.FromArgb(30, 64, 175) : Color.FromArgb(203, 213, 225), selected ? 2 : 1))
                e.Graphics.DrawRectangle(pen, tabRect.X + 1, tabRect.Y + 1, tabRect.Width - 2, tabRect.Height - 2);
        }

        // ---------- Scaffold dùng chung cho mọi tab ----------

        private DataGridView BuildScaffold(TabPage tab, string title, string standard,
            string condition, string note, out FlowLayoutPanel bar)
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 6, Padding = new Padding(12)
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 82));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 72));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            tab.Controls.Add(root);

            root.Controls.Add(MakeTitle(title), 0, 0);
            root.Controls.Add(MakeSubtitle(standard), 0, 1);
            root.Controls.Add(MakeCondition(condition), 0, 2);

            var box = new GroupBox
            {
                Dock = DockStyle.Fill, Text = "Tổ hợp kiểm tra", Padding = new Padding(10, 8, 16, 8)
            };
            root.Controls.Add(box, 0, 3);

            bar = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false, Margin = new Padding(0)
            };
            box.Controls.Add(bar);

            root.Controls.Add(MakeNote(note), 0, 4);

            var grid = CreateGrid();
            root.Controls.Add(grid, 0, 5);
            return grid;
        }

        // ---------- Factory tạo control ----------

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
            Text = text, Dock = DockStyle.Fill, AutoSize = false, Font = new Font("Arial", 10F),
            ForeColor = Color.DimGray, TextAlign = ContentAlignment.TopLeft,
            Padding = new Padding(2, 2, 2, 0)
        };

        private static Label MakeFieldLabel(string text, int width) => new Label
        {
            Text = text, AutoSize = false, Width = width, Height = CtrlHeight,
            TextAlign = ContentAlignment.MiddleLeft, Margin = new Padding(0, 6, 10, 0)
        };

        private static ComboBox MakeCombo(int width) => new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList, Width = width,
            Margin = new Padding(0, 6, 18, 0)
        };

        private static TextBox MakeTextBox(string value, int width) => new TextBox
        {
            Text = value, Width = width, Margin = new Padding(0, 7, 18, 0)
        };

        private static Button MakeButton(string text) => new Button
        {
            Text = text, Width = 112, Height = CtrlHeight, Margin = new Padding(0, 6, 10, 0)
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
                Margin = new Padding(0, 8, 0, 0),
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter },
                ColumnHeadersDefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
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
            if (fill) col.FillWeight = width;
            if (!string.IsNullOrWhiteSpace(format)) col.DefaultCellStyle.Format = format;
            grid.Columns.Add(col);
        }
    }
}
