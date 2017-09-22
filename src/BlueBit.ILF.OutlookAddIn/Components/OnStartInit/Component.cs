using BlueBit.ILF.OutlookAddIn.Common.Extensions;
using BlueBit.ILF.OutlookAddIn.Common.Utils;
using System;
using System.Windows.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BlueBit.ILF.OutlookAddIn.Components.OnStartInit
{
    [Obsolete("Hack for Outlook errors when first enumerate calendar folders")]
    public class Component :
        ISelfRegisteredComponent
    {
        public void Initialize(Outlook.Application app)
        {
            var timer = new DispatcherTimer();
            timer.Interval = new TimeSpan(0, 0, 5);
            timer.Tick += (s, e) =>
            {
                timer.Stop();

                Func<Outlook.Folder> getRootFolder = app
                    .GetNamespace("MAPI")
                    .GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar)
                    .As<Outlook.Folder>;

                using (var foldersSource = new FoldersSource(
                    getRootFolder(),
                    _ => true,
                    _ => true
                    ))
                    foldersSource.EnumFolders((fld, sel) => { });
            };
            timer.Start();
        }

        public void Execute()
        {
        }
    }
}
