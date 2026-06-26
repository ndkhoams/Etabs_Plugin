using ETABSv1;
using System;
using System.Windows.Forms;

namespace Etabs_Ultimate_Tools
{
    public class cPlugin
    {
        private cSapModel _sapModel;
        private cPluginCallback _callback;

        public void Main(ref cSapModel SapModel, ref cPluginCallback ISapPlugin)
        {
            _sapModel = SapModel;
            _callback = ISapPlugin;

            try
            {
                Application.EnableVisualStyles();
                using (var form = new ModelCheckForm(_sapModel))
                {
                    form.ShowDialog();
                }
                _callback.Finish(0);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Etabs Ultimate Tools", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _callback.Finish(1);
            }
        }
    }
}
