using BlueBit.ILF.OutlookAddIn.Diagnostics;
using BlueBit.ILF.OutlookAddIn.Properties;
using Microsoft.Office.Interop.Outlook;
using NLog;
using System.Linq;
using WinForms = System.Windows.Forms;

namespace BlueBit.ILF.OutlookAddIn.Components.OnSendEmailSizeChecker
{
    public class Component : ISelfRegisteredComponent
    {
        private static Logger _logger = LogManager.GetCurrentClassLogger();
        private IConfiguration _cfg;

        public Component(IConfiguration cfg)
        {
            _cfg = cfg;
        }

        public void Initialize(Application app)
        {
            app.ItemSend += App_ItemSend;
        }

        private void App_ItemSend(object Item, ref bool Cancel)
            => Cancel = _logger.OnEntryCall(() =>
            {
                var email = Item as MailItem;
                if (email == null) return false;
                if (email.Attachments.Count == 0) return false;
                var maxSize = _cfg.GetEmailSize();
                var size = email.Attachments.Cast<Attachment>().Sum(_ => _.Size);
                if (size <= maxSize) return false;
                var msg = string.Format(Resources.OnSendEmailSizeChecker_Message, maxSize);
                return WinForms.MessageBox.Show(msg, Resources.OnSendEmailSizeChecker_Caption, WinForms.MessageBoxButtons.YesNo) == WinForms.DialogResult.Yes;
            });
    }
}
