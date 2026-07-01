using ETABSv1;
using Strip_Modifier;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Strip_Rename
{
    public partial class Form1 : Form
    {
        private cSapModel _sapModel;
        private List<StripInfo> _preview;

        public Form1()
        {
            InitializeComponent();

            cboSortMode.DataSource = Enum.GetValues(typeof(SortMode));
            txtPrefix.Text = "S";
            numStart.Value = 1;
            numPad.Value = 3;
        }

        private void btnAttach_Click(object sender, EventArgs e)
        {
            try
            {
                var (_, sapModel) = EtabsConnector.AttachToRunningEtabs();
                _sapModel = sapModel;
                MessageBox.Show("Attach ETABS OK.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Attach error");
            }
        }

        private void btnPreview_Click(object sender, EventArgs e)
        {
            try
            {
                if (_sapModel == null) { MessageBox.Show("Chưa attach ETABS."); return; }

                var prefix = txtPrefix.Text?.Trim() ?? "";
                var startIndex = (int)numStart.Value;
                var padWidth = (int)numPad.Value;
                var mode = (SortMode)cboSortMode.SelectedItem;

                _preview = StripRenamerService.PreviewRename(_sapModel, prefix, startIndex, padWidth, mode);

                grid.DataSource = null;
                grid.DataSource = _preview;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Preview error");
            }
        }

        private void btnApply_Click(object sender, EventArgs e)
        {
            try
            {
                if (_sapModel == null) { MessageBox.Show("Chưa attach ETABS."); return; }
                if (_preview == null || _preview.Count == 0) { MessageBox.Show("Chưa preview."); return; }

                StripRenamerService.ApplyRename(_sapModel, _preview);
                MessageBox.Show("Rename OK.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Rename error");
            }
        }
    }
}