using System.Collections.Generic;

namespace BlueBit.ILF.OutlookAddIn.Components.OnAddAppointmentHandler
{
    public interface IConfiguration
    {
        IEnumerable<string> GetCalendarPrefixes();
        IEnumerable<string> GetDeafultCalendars();
    }
}
