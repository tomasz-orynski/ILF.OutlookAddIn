using BlueBit.ILF.OutlookAddIn.Common.Patterns;
using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BlueBit.ILF.OutlookAddIn.Common.Utils
{
    public interface IFolderSource
    {
        string ID { get; }
        string Name { get; }
        bool IsSelected { get; }

        IReadOnlyList<(string ID, string Name)> Categories { get; }
    }

    public interface IFoldersSource
    {
        IReadOnlyList<IFolderSource> Folders { get; }
        ICW<Outlook.Folder> GetFolder(IFolderSource folderSource);
    }

}
