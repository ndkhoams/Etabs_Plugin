using ETABSv1;
using System;
using System.Windows.Forms;

namespace Etabs_Ultimate_Tools
{
    public class cPlugin
    {
        private static ModelCheckForm _form;
        private static cPluginCallback _callback;

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
            _callback = ISapPlugin;
            ModelCheckForm form = null;

            try
            {
                Application.EnableVisualStyles();
                form = new ModelCheckForm(SapModel);
                _form = form;
                form.FormClosed += (_, _) =>
                {
                    if (ReferenceEquals(_form, form))
                    {
                        _form = null;
                        cPluginCallback callback = _callback;
                        _callback = null;
                        try { callback?.Finish(0); } catch { }
                    }
                };
                form.Show();
            }
            catch (Exception ex)
            {
                if (ReferenceEquals(_form, form))
                    _form = null;
                _callback = null;
                if (form != null && !form.IsDisposed)
                    form.Dispose();
                MessageBox.Show(ex.Message, "Etabs Ultimate Tools", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ISapPlugin.Finish(1);
            }
        }
    }
}
