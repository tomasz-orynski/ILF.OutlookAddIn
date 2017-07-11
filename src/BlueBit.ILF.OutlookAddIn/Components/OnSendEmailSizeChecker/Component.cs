using BlueBit.ILF.OutlookAddIn.Diagnostics;
using BlueBit.ILF.OutlookAddIn.Properties;
using Outlook = Microsoft.Office.Interop.Outlook;
using NLog;
using System.Linq;
using System.Windows;
using BlueBit.ILF.OutlookAddIn.Common.Extensions;

namespace BlueBit.ILF.OutlookAddIn.Components.OnSendEmailSizeChecker
{
    public class Component : ISelfRegisteredComponent
    {
        private static Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly IConfiguration _cfg;
        private Outlook.Application _app;

        public Component(IConfiguration cfg)
        {
            _cfg = cfg;
        }

        public void Initialize(Outlook.Application app)
        {
            _app = app;
            _app.ItemSend += OnItemSend;
        }

        private void OnItemSend(object item, ref bool cancel)
            => cancel = _logger.OnEntryCall(() =>
            {
                var email = item as Outlook.MailItem;
                if (email == null)
                    return false;
                if (email.Attachments.Count == 0)
                    return false;
                if (!email.Recipients.Cast<Outlook.Recipient>().Select(_ => _.Address).Any(_cfg.GetEmailGroups().AsEqualsFilter()))
                    return false;
                var maxSize = _cfg.GetEmailSize();
                if (maxSize < 0)
                    return false;
                var size = email.Attachments.Cast<Outlook.Attachment>().Sum(_ => _.Size);
                if (size <= maxSize)
                    return false;
                var msg = string.Format(Resources.OnSendEmailSizeChecker_Message, maxSize.ToStringWithSizeUnit());
                return MessageBox.Show(msg, Resources.OnSendEmailSizeChecker_Caption, MessageBoxButton.YesNo) == MessageBoxResult.Yes;
            });
    }
}
