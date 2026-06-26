using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Etabs_Ultimate_Tools
{
    public class ComboSelectForm : Form
    {
        private readonly CheckedListBox clbCombos;
        private readonly RadioButton rdoText;
        private readonly RadioButton rdoExcel;
        private int _lastClickedIndex = -1;

        public List<string> SelectedCombos { get; private set; } = new List<string>();
        public ExportFileType ExportType { get; private set; } = ExportFileType.Text;

        public ComboSelectForm(List<string> combos)
        {
            Text = "CSI Column Export ©2026.1 - KhoaND";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(474, 600);
            Font = new Font("Arial", 9F);

            Label lblTitle = new Label
            {
                Text = "Chọn tổ hợp tải (Load Combination):",
                Location = new Point(12, 12),
                AutoSize = true,
                Font = new Font("Arial", 9.5F, FontStyle.Bold)
            };
            Controls.Add(lblTitle);

            clbCombos = new CheckedListBox
            {
                Location = new Point(12, 38),
                Size = new Size(450, 410),
                CheckOnClick = true,
                IntegralHeight = false
            };
            clbCombos.MouseDown += ClbCombos_MouseDown;
            Controls.Add(clbCombos);

            foreach (string combo in combos)
            {
                clbCombos.Items.Add(combo, true);
            }

            Button btnSelectAll = new Button
            {
                Text = "Chọn tất cả",
                Location = new Point(12, 456),
                Size = new Size(150, 32)
            };
            btnSelectAll.Click += (s, e) => SetAllChecked(true);
            Controls.Add(btnSelectAll);

            Button btnDeselectAll = new Button
            {
                Text = "Bỏ chọn",
                Location = new Point(172, 456),
                Size = new Size(150, 32)
            };
            btnDeselectAll.Click += (s, e) => SetAllChecked(false);
            Controls.Add(btnDeselectAll);

            GroupBox grpFormat = new GroupBox
            {
                Text = "Định dạng xuất",
                Location = new Point(12, 496),
                Size = new Size(450, 56)
            };
            Controls.Add(grpFormat);

            rdoText = new RadioButton
            {
                Text = "Text (.txt)",
                Location = new Point(16, 22),
                AutoSize = true,
                Checked = true
            };
            grpFormat.Controls.Add(rdoText);

            rdoExcel = new RadioButton
            {
                Text = "Excel (.xlsx)",
                Location = new Point(180, 22),
                AutoSize = true
            };
            grpFormat.Controls.Add(rdoExcel);

            Button btnOK = new Button
            {
                Text = "Xuất",
                Location = new Point(250, 560),
                Size = new Size(100, 32),
                DialogResult = DialogResult.OK
            };
            btnOK.Click += BtnOK_Click;
            Controls.Add(btnOK);

            Button btnCancel = new Button
            {
                Text = "Hủy",
                Location = new Point(362, 560),
                Size = new Size(100, 32),
                DialogResult = DialogResult.Cancel
            };
            Controls.Add(btnCancel);

            AcceptButton = btnOK;
            CancelButton = btnCancel;
        }

        private void SetAllChecked(bool state)
        {
            for (int i = 0; i < clbCombos.Items.Count; i++)
                clbCombos.SetItemChecked(i, state);
        }

        private void ClbCombos_MouseDown(object sender, MouseEventArgs e)
        {
            int index = clbCombos.IndexFromPoint(e.Location);
            if (index < 0)
                return;

            if ((ModifierKeys & Keys.Shift) == Keys.Shift && _lastClickedIndex >= 0)
            {
                bool newState = clbCombos.GetItemChecked(_lastClickedIndex);
                int start = Math.Min(_lastClickedIndex, index);
                int end = Math.Max(_lastClickedIndex, index);

                for (int i = start; i <= end; i++)
                    clbCombos.SetItemChecked(i, newState);
            }
            else
            {
                _lastClickedIndex = index;
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            SelectedCombos = clbCombos.CheckedItems.Cast<string>().ToList();

            if (SelectedCombos.Count == 0)
            {
                MessageBox.Show("Vui lòng chọn ít nhất một tổ hợp.");
                DialogResult = DialogResult.None;
                return;
            }

            ExportType = rdoExcel.Checked ? ExportFileType.Excel : ExportFileType.Text;
            DialogResult = DialogResult.OK;
        }
    }
}
