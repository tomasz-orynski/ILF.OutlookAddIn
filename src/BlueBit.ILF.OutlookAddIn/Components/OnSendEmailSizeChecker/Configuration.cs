using System.Collections.Generic;

namespace BlueBit.ILF.OutlookAddIn.Components.OnSendEmailSizeChecker
{
    public interface IConfiguration
    {
        long GetEmailSize();
        IEnumerable<string> GetEmailGroups();
    }
}
