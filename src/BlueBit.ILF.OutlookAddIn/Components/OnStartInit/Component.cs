using BlueBit.ILF.OutlookAddIn.Common.Extensions;
using BlueBit.ILF.OutlookAddIn.Common.Utils;
using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BlueBit.ILF.OutlookAddIn.Components.OnStartInit
{
    [Obsolete("Hack for Outlook errors when first enumerate calendar folders")]
    public class Component :
        ISelfRegisteredComponent
    {
        public void Initialize(Outlook.Application app)
        {
            Func<Outlook.Folder> getRootFolder = app
                .GetNamespace("MAPI")
                .GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar)
                .As<Outlook.Folder>;

            using (var foldersSource = new FoldersSource(
                getRootFolder(),
                s => true,
                s => true
                ))
                foldersSource.EnumFolders((fld, sel) => { });
        }

        public void Execute()
        {
        }
    }
}
