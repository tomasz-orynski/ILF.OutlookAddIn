﻿using BlueBit.ILF.OutlookAddIn.Common.Patterns;
using BlueBit.ILF.OutlookAddIn.Common.Utils;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BlueBit.ILF.OutlookAddIn.Components
{
    public interface IEnviroment
    {
        string UserName { get; }
        ICW<Outlook.Application> Application { get; }
        ICW<Outlook.Folder> CalendarFolder { get; }
        ICW<Outlook.Items> CalendarItems { get; }

        IFoldersSource FoldersSource { get; }
    }
}
