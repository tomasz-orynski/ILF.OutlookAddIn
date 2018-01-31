using BlueBit.ILF.OutlookAddIn.Common.Patterns;
using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BlueBit.ILF.OutlookAddIn.Common.Utils
{
    public interface IFoldersSource
    {
        void EnumFolders(Action<ICW<Outlook.NavigationFolder>, bool> enumAction);
    }

}
