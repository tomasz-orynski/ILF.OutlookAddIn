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
    }

    public interface IFoldersSource
    {
        IReadOnlyDictionary<string, IFolderSource> Folders { get; }

        void OnFolders(IReadOnlyDictionary<string, IFolderSource> folders, Action<IFolderSource, ICW<Outlook.MAPIFolder>> action);
    }

    public static class FoldersSourceExtensions
    {
        public static void OnFolders(this IFoldersSource @this, IFolderSource folder, Action<IFolderSource, ICW<Outlook.MAPIFolder>> action)
            => @this.OnFolders(new Dictionary<string, IFolderSource>()
                {
                    [folder.ID] = folder,
                },
                action);

    }
}
