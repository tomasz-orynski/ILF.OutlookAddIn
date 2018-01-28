using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BlueBit.ILF.OutlookAddIn.Common.Utils
{
    public interface IFoldersSource
    {
        void EnumFolders(Action<Outlook.NavigationFolder, bool> enumAction);
    }

}
