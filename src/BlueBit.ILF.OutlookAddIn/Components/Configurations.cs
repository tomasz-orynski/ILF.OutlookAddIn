using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Win32;
using System.Collections.Generic;
using System;
using NLog;
using BlueBit.ILF.OutlookAddIn.Common.Patterns;

namespace BlueBit.ILF.OutlookAddIn.Components
{
    public class Configurations :
        IComponent,
        OnSendEmailSizeChecker.IConfiguration,
        OnAddAppointmentHandler.IConfiguration,
        SetDefaultCalendars.IConfiguration,
        OnStartInit.IConfiguration
    {
        private const string Path = @"HKEY_CURRENT_USER\Software\ILF\OutlookApp";
        private const string Calendar_Default = nameof(Calendar_Default);
        private const char Separator = ';';
        private static Logger _logger = LogManager.GetCurrentClassLogger();

        private static T GetValue<T>(string name, T defValue = default(T))
        {
            var value = Registry.GetValue(Path, name, null);
            if (value == null)
            {
                _logger.Warn($"Use default config value [{name}].");
                return defValue;
            }
            return (T)value;
        }
        private static void SetValue(string name, object value) => Registry.SetValue(Path, name, value);

        IEnumerable<string> OnSendEmailSizeChecker.IConfiguration.GetEmailGroups()
            => GetValue<string>("Email_groups", string.Empty).Split(Separator);

        long OnSendEmailSizeChecker.IConfiguration.GetEmailSize()
            => GetValue<int>("Email_size", -1) * 1024 * 1024;

        public IEnumerable<string> GetCalendarPrefixes()
            => GetValue<string>("Calendar_Prefix", string.Empty).Split(Separator);
        public IEnumerable<string> GetDeafultCalendars()
            => GetValue<string>(Calendar_Default, string.Empty).Split(Separator);
        public void SetDeafultCalendars(IEnumerable<string> calendars)
            => SetValue(Calendar_Default, string.Join(Separator.ToString(), calendars));
        public int GetInitOnStart()
            => GetValue<int>("InitOnStart", 10);
    }
}
