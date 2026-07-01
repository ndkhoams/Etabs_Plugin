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

        private class ModGroup
        {
            public string Title;
            public readonly List<TextBox> Boxes = new List<TextBox>();

            public double[] ReadValues()
            {
                var arr = new double[Boxes.Count];
                for (int i = 0; i < Boxes.Count; i++) arr[i] = ReadModValue(Boxes[i]);
                return arr;
            }
        }
    }
}
