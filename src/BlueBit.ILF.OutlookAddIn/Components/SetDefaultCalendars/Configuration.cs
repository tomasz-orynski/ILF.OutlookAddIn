using System.Collections.Generic;

namespace BlueBit.ILF.OutlookAddIn.Components.SetDefaultCalendars
{
    public interface IConfiguration
    {
        IEnumerable<string> GetCalendarPrefixes();
        IEnumerable<string> GetDeafultCalendars();
        void SetDeafultCalendars(IEnumerable<string> calendars);
    }
}
