using BlueBit.ILF.OutlookAddIn.Diagnostics;
using BlueBit.ILF.OutlookAddIn.Properties;
using Outlook = Microsoft.Office.Interop.Outlook;
using NLog;
using System.Linq;
using System.Windows;
using BlueBit.ILF.OutlookAddIn.Common.Extensions;

namespace BlueBit.ILF.OutlookAddIn.Components.OnSendEmailSizeChecker
{
    public class Component : IInitializedComponent
    {
        private static Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly IEnviroment _env;
        private readonly IConfiguration _cfg;

        public Component(
            IEnviroment env,
            IConfiguration cfg)
        {
            _env = env;
            _cfg = cfg;
        }

        public void Initialize()
        {
            _env.Application.Ref.ItemSend += OnItemSend;
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
