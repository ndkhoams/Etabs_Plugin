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
    }
}
