using BlueBit.ILF.OutlookAddIn.Common.Extensions;
using MoreLinq;
using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace BlueBit.ILF.OutlookAddIn.Components.OnAddAppointmentHandler
{
    class FoldersSource :
            IDisposable
    {
        private readonly Outlook.Application _application;
        private readonly IEnumerable<Tuple<Outlook.Folder, bool>> _foldersSource;
        private readonly IEnumerable<Action> _onDisposeActions;

        public FoldersSource(
            Outlook.Folder rootFolder,
            Func<string, bool> folderFilter,
            Func<string, bool> folderSelected
            )
        {
            Contract.Assert(rootFolder != null);
            Contract.Assert(folderFilter != null);
            Contract.Assert(folderSelected != null);

            var onDisposeActions = new List<Action>();
            _onDisposeActions = onDisposeActions;
            _application = rootFolder.Application;

            _foldersSource = (GetExplorer(rootFolder) ?? GetExplorer(rootFolder, onDisposeActions))
                .NavigationPane
                .Modules
                .GetNavigationModule(Outlook.OlNavigationModuleType.olModuleCalendar)
                .As<Outlook.CalendarModule>()
                .NavigationGroups
                .Cast<Outlook.NavigationGroup>()
                .SelectMany(_ => _.NavigationFolders.Cast<Outlook.NavigationFolder>())
                .Where(_ => _.Folder.FolderPath != rootFolder.FolderPath)
                .Select(_ => new { Folder = _.Folder.As<Outlook.Folder>(), _.DisplayName })
                .Where(_ => folderFilter(_.DisplayName))
                .Where(_=> CheckFolder(_.Folder))
                .Select(_ => Tuple.Create(_.Folder, folderSelected(_.DisplayName)));
        }

        public void Dispose()
        {
            _onDisposeActions.ForEach(_ => _.Invoke());
        }

        public void EnumFolders(Action<Outlook.Folder, bool> enumAction)
        {
            Contract.Assert(enumAction != null);
            _foldersSource
                .ForEach(_ => enumAction(_.Item1, _.Item2));
        }
        private Outlook.Explorer GetExplorer(Outlook.Folder folder)
            => _application.Explorers.Cast<Outlook.Explorer>().FirstOrDefault(_ => _.CurrentFolder.FolderPath == folder.FolderPath);

        private Outlook.Explorer GetExplorer(Outlook.Folder folder, IList<Action> onDisposeActions)
        {
            var explorer = folder.GetExplorer();
            onDisposeActions.Add(explorer.Close);
            return explorer;
        }

        static bool CheckFolder(Outlook.Folder folder)
        {
            try
            {
                var item = folder.Items.Add(Outlook.OlItemType.olAppointmentItem);
                item.Delete();
                return true;
            }
            catch
            {
            }
            return false;
        }
    }
}
