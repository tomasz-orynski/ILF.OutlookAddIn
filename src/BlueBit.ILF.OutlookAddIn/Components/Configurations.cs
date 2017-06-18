using Microsoft.Office.Interop.Outlook;
using Microsoft.Win32;
using System.Collections.Generic;
using System;

namespace BlueBit.ILF.OutlookAddIn.Components
{
    public class Configurations :
        ISelfRegisteredComponent,
        OnSendEmailSizeChecker.IConfiguration,
        OnAddAppointmentHandler.IConfiguration,
        SetDefaultCalendars.IConfiguration
    {
        private const string Path = @"HKEY_CURRENT_USER\Software\ILF\OutlookApp";
        private const string Calendar_Default = nameof(Calendar_Default);
        private const char Separator = ';';

        private static T GetValue<T>(string name) => (T)Registry.GetValue(Path, name, null);
        private static void SetValue(string name, object value) => Registry.SetValue(Path, name, value);

        IEnumerable<string> OnSendEmailSizeChecker.IConfiguration.GetEmailGroups()
            => GetValue<string>("Email_groups").Split(Separator);

        long OnSendEmailSizeChecker.IConfiguration.GetEmailSize()
            => GetValue<int>("Email_size") * 1024 * 1024;

        public IEnumerable<string> GetCalendarPrefixes()
            => GetValue<string>("Calendar_Prefix").Split(Separator);
        public IEnumerable<string> GetDeafultCalendars()
            => GetValue<string>(Calendar_Default).Split(Separator);
        public void SetDeafultCalendars(IEnumerable<string> calendars)
            => SetValue(Calendar_Default, string.Join(Separator.ToString(), calendars));

        public void Initialize(Application app)
        {
            //TODO-TO: check that configuration exists...
        }

    }
}
