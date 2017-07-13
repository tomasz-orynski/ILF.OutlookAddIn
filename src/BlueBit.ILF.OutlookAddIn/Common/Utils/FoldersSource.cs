using BlueBit.ILF.OutlookAddIn.Common.Extensions;
using BlueBit.ILF.OutlookAddIn.Diagnostics;
using MoreLinq;
using NLog;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.Contracts;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BlueBit.ILF.OutlookAddIn.Common.Utils
{
    class FoldersSource :
            IDisposable
    {
        private static Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly Outlook.Application _application;
        private readonly IEnumerable<Tuple<Outlook.NavigationFolder, bool>> _foldersSource;
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
                .SafeWhere(_ => folderFilter(_.DisplayName))
                .SafeWhere(_ => _.Folder.FolderPath != rootFolder.FolderPath)
                .SafeWhere(_ => CheckFolder(_.Folder.As<Outlook.Folder>()))
                .OrderBy(_ => _.DisplayName)
                .Select(_ => Tuple.Create(_, folderSelected(_.DisplayName)))
                .SafeToList()
                ;
        }

        public void Dispose()
        {
            _onDisposeActions.ForEach(_ => _.Invoke());
        }

        public void EnumFolders(Action<Outlook.NavigationFolder, bool> enumAction)
        {
            Contract.Assert(enumAction != null);
            _foldersSource
                .ForEach(_ => enumAction(_.Item1, _.Item2));
        }

        public IEnumerable<Outlook.Folder> GetFolders()
            => _foldersSource.Select(_ => _.Item1.Folder.As<Outlook.Folder>());

        private Outlook.Explorer GetExplorer(Outlook.Folder folder)
            => _application.Explorers.Cast<Outlook.Explorer>().FirstOrDefault(_ => _.CurrentFolder.FolderPath == folder.FolderPath);

        private Outlook.Explorer GetExplorer(Outlook.Folder folder, IList<Action> onDisposeActions)
        {
            var explorer = folder.GetExplorer();
            //onDisposeActions.Add(explorer.Close);
            return explorer;
        }

        static bool CheckFolder(Outlook.Folder folder)
        {
            try
            {
                var item = (Outlook.AppointmentItem)folder.Items.Add(Outlook.OlItemType.olAppointmentItem);
                item.Delete();
                return true;
            }
            catch(Exception e)
            {
                var msg = $"{nameof(CheckFolder)}({folder.Name})";
                _logger.Warn(e, msg);
                if (Debugger.IsAttached)
                    Debugger.Break();
            }
            return false;
        }
    }
}
