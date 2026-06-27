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

        // ---------- Tab xuất nội lực cột (CSI Column) ----------

        private const string ColTitle = "Xuất nội lực cột";

        private void BuildColumnExportTab(TabPage tab)
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(12)
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            tab.Controls.Add(root);

            root.Controls.Add(MakeTitle("XUẤT NỘI LỰC CỘT / VÁCH"), 0, 0);
            root.Controls.Add(MakeSubtitle("(Xuất nội lực theo định dạng của CSI Column và Prokon, đơn vị kN-m)"), 0, 1);

            var main = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Margin = new Padding(0, 6, 0, 0)
            };
            main.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 380));
            main.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.Controls.Add(main, 0, 2);

            // ----- Cột trái: chọn tổ hợp + tùy chọn xuất -----
            var left = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 5, Margin = new Padding(0, 0, 10, 0)
            };
            left.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
            left.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            left.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            left.RowStyles.Add(new RowStyle(SizeType.Absolute, 56));
            left.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
            main.Controls.Add(left, 0, 0);

            left.Controls.Add(new Label
            {
                Text = "Chọn Load Combination:",
                Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft
            }, 0, 0);

            clbColCombos = new CheckedListBox
            {
                Dock = DockStyle.Fill, CheckOnClick = true, IntegralHeight = false,
                BorderStyle = BorderStyle.FixedSingle
            };
            clbColCombos.MouseDown += ClbColCombos_MouseDown;
            left.Controls.Add(clbColCombos, 0, 1);

            var selBar = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = false,
                Margin = new Padding(0, 6, 0, 0)
            };
            left.Controls.Add(selBar, 0, 2);

            var btnSelAll = new Button { Text = "Chọn tất cả", Width = 104, Height = 28, Margin = new Padding(0, 0, 6, 0) };
            btnSelAll.Click += (s, e) => SetColCombosChecked(true);
            selBar.Controls.Add(btnSelAll);

            var btnDeselAll = new Button { Text = "Bỏ chọn", Width = 104, Height = 28, Margin = new Padding(0, 0, 6, 0) };
            btnDeselAll.Click += (s, e) => SetColCombosChecked(false);
            selBar.Controls.Add(btnDeselAll);

            btnColPreview = new Button { Text = "Xem trước", Width = 104, Height = 28, Margin = new Padding(0, 0, 0, 0) };
            btnColPreview.Click += (s, e) => PreviewColumnForces();
            selBar.Controls.Add(btnColPreview);

            var fmtRow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = false,
                Margin = new Padding(0, 6, 0, 0)
            };
            left.Controls.Add(fmtRow, 0, 3);

            btnColExportText = new Button { Text = "Xuất Text (.txt)", Width = 168, Height = 42, Enabled = false, Margin = new Padding(0, 3, 12, 0) };
            btnColExportText.Click += (s, e) => ExportColumnForces(false);
            fmtRow.Controls.Add(btnColExportText);

            btnColExportExcel = new Button { Text = "Xuất Excel (.xlsx)", Width = 168, Height = 42, Enabled = false, Margin = new Padding(0, 3, 0, 0) };
            btnColExportExcel.Click += (s, e) => ExportColumnForces(true);
            fmtRow.Controls.Add(btnColExportExcel);

            lblColInfo = new Label
            {
                Dock = DockStyle.Fill, ForeColor = Color.DimGray, TextAlign = ContentAlignment.MiddleLeft
            };
            left.Controls.Add(lblColInfo, 0, 4);

            var right = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2
            };
            right.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
            right.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            main.Controls.Add(right, 1, 0);

            right.Controls.Add(new Label
            {
                Text = "Preview:",
                Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft
            }, 0, 0);

            dgvColPreview = CreateGrid();
            right.Controls.Add(dgvColPreview, 0, 1);

            AddColumnExportGridColumns();

            lblColInfo.Text = "Chọn cột/vách trong ETABS trước khi mở tool, chọn tổ hợp rồi bấm Xem trước.";
        }

        private void ClbColCombos_MouseDown(object sender, MouseEventArgs e)
        {
            int index = clbColCombos.IndexFromPoint(e.Location);
            if (index < 0) return;

            if ((Control.ModifierKeys & Keys.Shift) == Keys.Shift && _lastColIndex >= 0 && _lastColIndex != index)
            {
                bool target = !clbColCombos.GetItemChecked(index);
                int start = Math.Min(_lastColIndex, index);
                int end = Math.Max(_lastColIndex, index);
                for (int i = start; i <= end; i++)
                {
                    if (i == index) continue;
                    clbColCombos.SetItemChecked(i, target);
                }
            }

            _lastColIndex = index;
        }

        private void SetColCombosChecked(bool state)
        {
            for (int i = 0; i < clbColCombos.Items.Count; i++)
                clbColCombos.SetItemChecked(i, state);
        }

        private void AddColumnExportGridColumns()
        {
            dgvColPreview.Columns.Clear();
            dgvColPreview.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            AddColumn(dgvColPreview, "Name", "NAME", 240, null, true);
            AddColumn(dgvColPreview, "PU", "PU", 80, "0.##", true);
            AddColumn(dgvColPreview, "MUXT", "MUXT", 80, "0.##", true);
            AddColumn(dgvColPreview, "MUYT", "MUYT", 80, "0.##", true);
            AddColumn(dgvColPreview, "MUXB", "MUXB", 80, "0.##", true);
            AddColumn(dgvColPreview, "MUYB", "MUYB", 80, "0.##", true);
        }

        private void PreviewColumnForces()
        {
            var combos = clbColCombos.CheckedItems.Cast<string>().ToList();
            if (combos.Count == 0)
            {
                Warn("Chưa chọn tổ hợp tải nào.", ColTitle);
                return;
            }

            int cols, piers;
            try
            {
                _colRows = ColumnForceExporter.Compute(_sap, combos, out cols, out piers);
            }
            catch (Exception ex)
            {
                Warn(ex.Message, ColTitle);
                return;
            }

            if (cols == 0 && piers == 0)
            {
                _colRows = new List<ForceRow>();
                dgvColPreview.DataSource = null;
                lblColInfo.Text = "Chưa chọn cột hoặc vách (Pier) nào trong ETABS.";
                btnColExportText.Enabled = false;
                btnColExportExcel.Enabled = false;
                Warn("Chưa chọn cột hoặc vách (Pier) nào trong ETABS.", ColTitle);
                return;
            }

            dgvColPreview.DataSource = null;
            dgvColPreview.DataSource = _colRows;

            lblColInfo.Text = "Cột: " + cols + "  |  Vách: " + piers + "  |  Dòng: " + _colRows.Count;

            if (_colRows.Count == 0)
                Warn("Không có nội lực để xuất. Hãy chạy Analyze trước.", ColTitle);

            bool hasData = _colRows.Count > 0;
            btnColExportText.Enabled = hasData;
            btnColExportExcel.Enabled = hasData;
        }

        private void ExportColumnForces(bool excel)
        {
            if (_colRows == null || _colRows.Count == 0)
            {
                Warn("Chưa có dữ liệu. Hãy bấm Xem trước trước khi xuất.", ColTitle);
                return;
            }

            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = excel ? "Excel file (*.xlsx)|*.xlsx" : "Text File (*.txt)|*.txt";
                sfd.FileName = excel ? "Column_Forces.xlsx" : "Column_Forces.txt";
                if (sfd.ShowDialog() != DialogResult.OK) return;

                try
                {
                    if (excel)
                        ColumnForceExporter.WriteExcel(_colRows, sfd.FileName);
                    else
                        ColumnForceExporter.WriteText(_colRows, sfd.FileName);
                }
                catch (Exception ex)
                {
                    Warn(ex.Message, ColTitle);
                    return;
                }

                Info("Đã xuất: " + sfd.FileName, ColTitle);
            }
        }

        // ---------- Tab Property Modifiers ----------

        private const string ModTitle = "Property Modifiers";

        // Thứ tự hàng phải khớp với mảng modifier của ETABS:
        //  Frame (8): Area, As2, As3, J, I22, I33, Mass, Weight
        //  Area (10): F11, F22, F12, M11, M22, M12, V13, V23, Mass, Weight
        private static readonly (string Label, string Def)[] BeamRows =
        {
            ("Area", "1.00"), ("As2", "1.00"), ("As3", "1.00"), ("J", "0.10"),
            ("I22", "0.50"), ("I33", "0.50"), ("Mass", "1.00"), ("Weight", "1.00")
        };

        private static readonly (string Label, string Def)[] ColRows =
        {
            ("Area", "1.00"), ("As2", "1.00"), ("As3", "1.00"), ("J", "0.50"),
            ("I22", "0.50"), ("I33", "0.50"), ("Mass", "1.00"), ("Weight", "1.00")
        };

        private static readonly (string Label, string Def)[] SlabRows =
        {
            ("F11", "1.00"), ("F22", "1.00"), ("F12", "1.00"), ("M11", "0.50"), ("M22", "0.50"),
            ("M12", "0.50"), ("V13", "1.00"), ("V23", "1.00"), ("Mass", "1.00"), ("Weight", "1.00")
        };

        private static readonly (string Label, string Def)[] WallRows =
        {
            ("F11", "0.50"), ("F22", "0.50"), ("F12", "0.50"), ("M11", "0.50"), ("M22", "0.50"),
            ("M12", "0.50"), ("V13", "0.50"), ("V23", "0.50"), ("Mass", "1.00"), ("Weight", "1.00")
        };

        private void BuildModifierTab(TabPage tab)
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 5, Padding = new Padding(12)
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 56));
            tab.Controls.Add(root);

            root.Controls.Add(MakeTitle("PROPERTY MODIFIERS"), 0, 0);
            root.Controls.Add(MakeSubtitle("(Gán hệ số tiết diện cho cấu kiện đang chọn trong ETABS)"), 0, 1);

            var groups = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 4, RowCount = 1, Margin = new Padding(0, 6, 0, 0)
            };
            for (int i = 0; i < 4; i++) groups.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25f));
            root.Controls.Add(groups, 0, 3);

            _modBeam = BuildModGroup(groups, 0, "Beam", BeamRows);
            _modCol = BuildModGroup(groups, 1, "Column", ColRows);
            _modSlab = BuildModGroup(groups, 2, "Slab", SlabRows);
            _modWall = BuildModGroup(groups, 3, "Wall", WallRows);

            var bar = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = false,
                Margin = new Padding(0, 8, 0, 0)
            };
            root.Controls.Add(bar, 0, 4);

            btnModApply = new Button { Text = "Apply Modifiers", Width = 170, Height = 40, Margin = new Padding(0, 0, 10, 0) };
            btnModApply.Click += (s, e) => ApplyModifiers();
            bar.Controls.Add(btnModApply);

            btnModRollback = new Button { Text = "Reset (= 1.0)", Width = 140, Height = 40, Margin = new Padding(0, 0, 10, 0) };
            btnModRollback.Click += (s, e) => RollbackModifiers();
            bar.Controls.Add(btnModRollback);

            lblModInfo = new Label
            {
                AutoSize = false, Width = 620, Height = 40, ForeColor = Color.DimGray,
                TextAlign = ContentAlignment.MiddleLeft, Margin = new Padding(8, 0, 0, 0)
            };
            lblModInfo.Text = "Chọn cấu kiện trong ETABS trước khi mở tool, bấm Apply Modifiers để gán hệ số.";
            bar.Controls.Add(lblModInfo);
        }

        private ModGroup BuildModGroup(TableLayoutPanel parent, int col, string title, (string Label, string Def)[] rows)
        {
            var g = new GroupBox
            {
                Text = title, Dock = DockStyle.Fill, Padding = new Padding(8, 6, 8, 6), Margin = new Padding(0, 0, 10, 0)
            };
            var t = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 2, RowCount = 0, AutoSize = true };
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 64));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            g.Controls.Add(t);

            var group = new ModGroup { Title = title };
            foreach (var r in rows)
            {
                int row = t.RowCount;
                t.RowCount = row + 1;
                t.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
                t.Controls.Add(new Label
                {
                    Text = r.Label, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft
                }, 0, row);
                var txt = new TextBox { Text = r.Def, Dock = DockStyle.Fill, Margin = new Padding(2, 4, 2, 4) };
                t.Controls.Add(txt, 1, row);
                group.Boxes.Add(txt);
            }

            parent.Controls.Add(g, col, 0);
            return group;
        }

        private void ApplyModifiers()
        {
            int numberItems = 0;
            int[] objectTypes = null;
            string[] objectNames = null;
            _sap.SelectObj.GetSelected(ref numberItems, ref objectTypes, ref objectNames);

            if (numberItems == 0)
            {
                Warn("Chưa chọn cấu kiện nào.", ModTitle);
                return;
            }

            double[] beam, col, slab, wall;
            try
            {
                beam = _modBeam.ReadValues();
                col = _modCol.ReadValues();
                slab = _modSlab.ReadValues();
                wall = _modWall.ReadValues();
            }
            catch
            {
                Warn("Giá trị nhập không hợp lệ. Hãy nhập số, ví dụ 0.50", ModTitle);
                return;
            }

            int cBeam = 0, cCol = 0, cSlab = 0, cWall = 0;
            for (int i = 0; i < numberItems; i++)
            {
                string name = objectNames[i];
                if (objectTypes[i] == 2)
                {
                    if (IsColumnFrame(name)) { var m = (double[])col.Clone(); _sap.FrameObj.SetModifiers(name, ref m); cCol++; }
                    else { var m = (double[])beam.Clone(); _sap.FrameObj.SetModifiers(name, ref m); cBeam++; }
                }
                else if (objectTypes[i] == 5)
                {
                    if (IsWallArea(name)) { var m = (double[])wall.Clone(); _sap.AreaObj.SetModifiers(name, ref m); cWall++; }
                    else { var m = (double[])slab.Clone(); _sap.AreaObj.SetModifiers(name, ref m); cSlab++; }
                }
            }

            _sap.View.RefreshView(0, false);

            lblModInfo.Text = "Đã gán — Dầm: " + cBeam + "  |  Cột: " + cCol + "  |  Sàn: " + cSlab + "  |  Vách: " + cWall;

            Info("Đã gán modifier:\n- Dầm: " + cBeam + "\n- Cột: " + cCol + "\n- Sàn: " + cSlab + "\n- Vách: " + cWall, ModTitle);
        }

        private void RollbackModifiers()
        {
            int numberItems = 0;
            int[] objectTypes = null;
            string[] objectNames = null;
            _sap.SelectObj.GetSelected(ref numberItems, ref objectTypes, ref objectNames);

            if (numberItems == 0)
            {
                Warn("Chưa chọn cấu kiện nào để reset.", ModTitle);
                return;
            }

            int cFrame = 0, cArea = 0;
            for (int i = 0; i < numberItems; i++)
            {
                string name = objectNames[i];
                if (objectTypes[i] == 2)
                {
                    var m = Ones(8);
                    _sap.FrameObj.SetModifiers(name, ref m); cFrame++;
                }
                else if (objectTypes[i] == 5)
                {
                    var m = Ones(10);
                    _sap.AreaObj.SetModifiers(name, ref m); cArea++;
                }
            }

            _sap.View.RefreshView(0, false);

            lblModInfo.Text = "Đã reset về 1.0 — Frame: " + cFrame + "  |  Area: " + cArea;

            Info("Đã reset về 1.0:\n- Frame: " + cFrame + "\n- Area: " + cArea, ModTitle);
        }

        private static double[] Ones(int n)
        {
            var m = new double[n];
            for (int i = 0; i < n; i++) m[i] = 1.0;
            return m;
        }

        private static double ReadModValue(TextBox t)
        {
            string v = t.Text.Trim().Replace(",", ".");
            return double.Parse(v, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
        }

        private bool IsColumnFrame(string frameName)
        {
            string p1 = "", p2 = "";
            _sap.FrameObj.GetPoints(frameName, ref p1, ref p2);
            double xi = 0, yi = 0, zi = 0, xj = 0, yj = 0, zj = 0;
            _sap.PointObj.GetCoordCartesian(p1, ref xi, ref yi, ref zi);
            _sap.PointObj.GetCoordCartesian(p2, ref xj, ref yj, ref zj);
            double dx = Math.Abs(xj - xi), dy = Math.Abs(yj - yi), dz = Math.Abs(zj - zi);
            double horizontal = Math.Sqrt(dx * dx + dy * dy);
            return dz > horizontal;
        }

        private bool IsWallArea(string areaName)
        {
            int n = 0;
            string[] pts = null;
            _sap.AreaObj.GetPoints(areaName, ref n, ref pts);
            double zMin = double.MaxValue, zMax = double.MinValue;
            for (int i = 0; i < n; i++)
            {
                double x = 0, y = 0, z = 0;
                _sap.PointObj.GetCoordCartesian(pts[i], ref x, ref y, ref z);
                zMin = Math.Min(zMin, z);
                zMax = Math.Max(zMax, z);
            }
            return Math.Abs(zMax - zMin) > 0.10;
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

        // ---------- Các tab ----------

        private void BuildPDeltaTab(TabPage tab)
        {
            dgv = BuildScaffold(tab,
                "KIỂM TRA ĐIỀU KIỆN P-DELTA",
                "(Theo TCVN 9386-1:2025)",
                "θ = dr / h × Ptot / Vtot =  q × drift × Ptot / Vtot (mục 4.4.2.2 eq. 4.28)",
                "drift = Δ/h được xác định từ hệ quả của tác động động đất thiết kế (mục 4.3.4); Vtot là tổng lực cắt tầng do động đất gây ra; 2 thành phần này đều lấy từ tổ hợp các thành phần phương ngang của động đất SRSS(EX;EY) (mục 4.3.3.5.1.2b). Ptot tự động lấy từ Mass Summary by Story.",
                out var bar);

            bar.Controls.Add(MakeFieldLabel("Tổ hợp:", 68));
            cboCombo = MakeCombo(240); bar.Controls.Add(cboCombo);
            bar.Controls.Add(MakeFieldLabel("q:", 22));
            txtQ = MakeTextBox("1.5", 60); bar.Controls.Add(txtQ);

            btnRun = MakeButton("Tính toán"); btnRun.Click += (s, e) => RunCheck(); bar.Controls.Add(btnRun);
            btnExport = MakeButton("Xuất Excel"); btnExport.Enabled = false; btnExport.Click += (s, e) => ExportPDelta(); bar.Controls.Add(btnExport);
            btnClose = MakeButton("Đóng"); btnClose.Width = 84; btnClose.Click += (s, e) => Close(); bar.Controls.Add(btnClose);

            AddPDeltaGridColumns();
        }

        private void BuildWindTab(TabPage tab)
        {
            dgvWind = BuildScaffold(tab,
                "KIỂM TRA CHUYỂN VỊ ĐỈNH CÔNG TRÌNH",
                "(Theo TCVN 2737:2023)",
                "Điều kiện kiểm tra: f ≤ fu",
                "Giới hạn chuyển vị ngang tổng thể là H/500. H được tính là khoảng cách từ mặt móng đến mái.",
                out var bar);

            bar.Controls.Add(MakeFieldLabel("Tổ hợp gió:", 78));
            cboWindCombo = MakeCombo(240); bar.Controls.Add(cboWindCombo);

            btnWindRun = MakeButton("Tính toán"); btnWindRun.Click += (s, e) => RunWindCheck(); bar.Controls.Add(btnWindRun);
            btnWindExport = MakeButton("Xuất Excel"); btnWindExport.Enabled = false; btnWindExport.Click += (s, e) => ExportWind(); bar.Controls.Add(btnWindExport);

            AddWindGridColumns();
        }

        private void BuildWindDriftTab(TabPage tab)
        {
            dgvWindDrift = BuildScaffold(tab,
                "KIỂM TRA CHUYỂN VỊ LỆCH TẦNG DO TẢI TRỌNG GIÓ",
                "(Theo TCVN 2737:2023)",
                "Điều kiện: drift = Δ/h ≤ 1/500 cho từng tầng",
                "Drift lấy trực tiếp từ ETABS Story Drifts theo tổ hợp gió.",
                out var bar);

            bar.Controls.Add(MakeFieldLabel("Tổ hợp gió:", 78));
            cboWindDriftCombo = MakeCombo(240); bar.Controls.Add(cboWindDriftCombo);

            btnWindDriftRun = MakeButton("Tính toán"); btnWindDriftRun.Click += (s, e) => RunWindDriftCheck(); bar.Controls.Add(btnWindDriftRun);
            btnWindDriftExport = MakeButton("Xuất Excel"); btnWindDriftExport.Enabled = false; btnWindDriftExport.Click += (s, e) => ExportWindDrift(); bar.Controls.Add(btnWindDriftExport);

            AddWindDriftGridColumns();
        }

        private void BuildSeismicDriftTab(TabPage tab)
        {
            dgvSeis = BuildScaffold(tab,
                "KIỂM TRA CHUYỂN VỊ LỆCH TẦNG DO TẢI TRỌNG ĐỘNG ĐẤT",
                "(Theo TCVN 9386-1:2025)",
                "Điều kiện hạn chế hư hỏng: dr·ν ≤ limit·h  ⇔  drift ≤ limit/(ν·q) (mục 4.4.3.2)",
                "drift = de/h (đàn hồi) lấy từ ETABS Story Drifts. dr = q × de là chuyển vị ngang thiết kế tương đối giữa các tầng. Drift lấy từ tổ hợp các thành phần phương ngang của động đất SRSS(EX;EY)." +
                "\nCHÚ THÍCH: Các giá trị khác nhau của ν phụ thuộc vào các nguy cơ động đất và vào cấp hậu quả của công trình, khuyến nghị như sau: ν = 0,4 cho các cấp hậu quả C3-a và C3-b, và ν = 0,5 cho các cấp hậu quả C1 và C2.",
                out var bar);

            bar.Controls.Add(MakeFieldLabel("Tổ hợp động đất:", 110));
            cboSeisCombo = MakeCombo(220); bar.Controls.Add(cboSeisCombo);
            bar.Controls.Add(MakeFieldLabel("q:", 22));
            txtSeisQ = MakeTextBox("1.5", 50); bar.Controls.Add(txtSeisQ);
            bar.Controls.Add(MakeFieldLabel("ν:", 22));
            txtSeisNu = MakeTextBox("0.4", 50); bar.Controls.Add(txtSeisNu);
            bar.Controls.Add(MakeFieldLabel("limit:", 38));
            cboSeisLimit = MakeCombo(150);
            cboSeisLimit.Items.AddRange(new object[] { "0.005 (giòn)", "0.0075 (dẻo)", "0.010 (không cản trở)" });
            cboSeisLimit.SelectedIndex = 0;
            bar.Controls.Add(cboSeisLimit);

            btnSeisRun = MakeButton("Tính toán"); btnSeisRun.Click += (s, e) => RunSeismicDriftCheck(); bar.Controls.Add(btnSeisRun);
            btnSeisExport = MakeButton("Xuất Excel"); btnSeisExport.Enabled = false; btnSeisExport.Click += (s, e) => ExportSeismic(); bar.Controls.Add(btnSeisExport);

            AddSeismicDriftGridColumns();
        }

        private void BuildAxialTab(TabPage tab)
        {
            dgvAxial = BuildScaffold(tab,
                "KIỂM TRA HỆ SỐ LỰC DỌC QUY ĐỔI",
                "(Theo TCVN 9386-1:2025)",
                "ʊd = Ned/(Ac·fcd) ≤ 0.65 (cột) / 0.40 (vách)  |  fcd = αcc·fck/γc",
                "Chọn cột hoặc vách (Pier) trong ETABS trước khi mở tool.",
                out var bar);

            bar.Controls.Add(MakeFieldLabel("Bê tông:", 88));
            cboAxialConcrete = MakeCombo(110);
            cboAxialConcrete.Items.AddRange(new object[] { "B15", "B20", "B22.5", "B25", "B30", "B35", "B40", "B45", "B50", "B55", "B60", "B70", "B80" });
            cboAxialConcrete.SelectedItem = "B30";
            if (cboAxialConcrete.SelectedIndex < 0 && cboAxialConcrete.Items.Count > 0) cboAxialConcrete.SelectedIndex = 0;
            bar.Controls.Add(cboAxialConcrete);

            bar.Controls.Add(MakeFieldLabel("Combo:", 56));
            cboAxialCombo = MakeCombo(220); bar.Controls.Add(cboAxialCombo);

            btnAxialRun = MakeButton("Tính toán"); btnAxialRun.Click += (s, e) => RunAxialCheck(); bar.Controls.Add(btnAxialRun);
            btnAxialExport = MakeButton("Xuất Excel"); btnAxialExport.Enabled = false; btnAxialExport.Click += (s, e) => ExportAxial(); bar.Controls.Add(btnAxialExport);

            lblAxialInfo = new Label
            {
                AutoSize = false, Width = 280, Height = CtrlHeight,
                TextAlign = ContentAlignment.MiddleLeft, ForeColor = Color.DimGray,
                Margin = new Padding(8, 4, 0, 0)
            };
            bar.Controls.Add(lblAxialInfo);

            AddAxialGridColumns();
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
            if (fill) col.FillWeight = width;
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
            AddColumn(dgvWindDrift, "Height", "h tầng (m)