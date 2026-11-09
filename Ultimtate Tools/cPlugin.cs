using ETABSv1;
using System;
using System.Windows.Forms;

namespace Etabs_Ultimate_Tools
{
    public class cPlugin
    {
        private static ModelCheckForm _form;

        public void Main(ref cSapModel SapModel, ref cPluginCallback ISapPlugin)
        {
            if (_form != null && !_form.IsDisposed)
            {
                if (_form.WindowState == FormWindowState.Minimized)
                    _form.WindowState = FormWindowState.Normal;
                _form.BringToFront();
                _form.Activate();
                return;
            }

            _form = null;
            ModelCheckForm form = null;

            try
            {
                Application.EnableVisualStyles();
                form = new ModelCheckForm(SapModel);
                _form = form;
                form.FormClosed += (_, _) =>
                {
                    if (ReferenceEquals(_form, form))
                        _form = null;
                };
                form.Show();
            }
            catch (Exception ex)
            {
                if (ReferenceEquals(_form, form))
                    _form = null;
                if (form != null && !form.IsDisposed)
                    form.Dispose();
                MessageBox.Show(ex.Message, "Etabs Ultimate Tools", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ISapPlugin.Finish(1);
            }
        }
    }
}
