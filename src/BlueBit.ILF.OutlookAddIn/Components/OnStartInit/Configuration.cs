using System.Collections.Generic;

namespace BlueBit.ILF.OutlookAddIn.Components.OnStartInit
{
    public interface IConfiguration : 
        SetDefaultCalendars.IConfiguration
    {
        bool GetInitOnStart();
    }
}
