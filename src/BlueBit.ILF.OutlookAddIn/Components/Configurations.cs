using Microsoft.Office.Interop.Outlook;
using Microsoft.Win32;
using System.Collections.Generic;

namespace BlueBit.ILF.OutlookAddIn.Components
{
    public class Configurations :
        ISelfRegisteredComponent,
        OnSendEmailSizeChecker.IConfiguration
    {
        private const string Path = @"HKEY_CURRENT_USER\Software\ILF\OutlookApp";
        private const char Separator = ';';

        private static T GetValue<T>(string name) => (T)Registry.GetValue(Path, name, null);

        IEnumerable<string> OnSendEmailSizeChecker.IConfiguration.GetEmailGroups()
            => GetValue<string>("Email_groups").Split(Separator);

        long OnSendEmailSizeChecker.IConfiguration.GetEmailSize()
            => GetValue<int>("Email_size") * 1024 * 1024;

        public void Initialize(Application app)
        {
            //TODO-TO: check that configuration exists...
        }
    }
}
