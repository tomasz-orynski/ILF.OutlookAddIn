using BlueBit.ILF.OutlookAddIn.Common.Extensions;
using BlueBit.ILF.OutlookAddIn.Common.Utils;
using BlueBit.ILF.OutlookAddIn.Diagnostics;
using MoreLinq;
using NLog;
using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BlueBit.ILF.OutlookAddIn.Components.OnStartInit
{
    partial class Component
    {
        private class _FoldersSource : IFoldersSource
        {
            private static Logger _logger = LogManager.GetCurrentClassLogger();
            private readonly Outlook.Application _application;
            private readonly IEnumerable<Tuple<Outlook.NavigationFolder, bool>> _foldersSource;

            public _FoldersSource(
                Outlook.Folder rootFolder,
                Func<string, bool> folderFilter,
                Func<string, bool> folderSelected
                )
            {
                Contract.Assert(rootFolder != null);
                Contract.Assert(folderFilter != null);
                Contract.Assert(folderSelected != null);

                _application = rootFolder.Application;
                _foldersSource = (GetExplorer(rootFolder.FolderPath) ?? GetExplorer(rootFolder))
                    .NavigationPane
                    .Modules
                    .GetNavigationModule(Outlook.OlNavigationModuleType.olModuleCalendar)
                    .As<Outlook.CalendarModule>()
                    .NavigationGroups
                    .Cast<Outlook.NavigationGroup>()
                    .SelectMany(_ => _.NavigationFolders.Cast<Outlook.NavigationFolder>())
                    .SafeWhere(_ => folderFilter(_.DisplayName))
                    .SafeWhere(_ => _.Folder.FolderPath != rootFolder.FolderPath)
                    .Select(_ => Tuple.Create(_, folderSelected(_.DisplayName)))
                    .SafeToList()
                    ;
            }

            public void EnumFolders(Action<Outlook.NavigationFolder, bool> enumAction)
                => _logger.OnEntryCall(() =>
                {
                    Contract.Assert(enumAction != null);
                    _foldersSource
                        .ForEach(_ => enumAction(_.Item1, _.Item2));
                });

            private Outlook.Explorer GetExplorer(string folderPath)
                => _application.Explorers.Cast<Outlook.Explorer>().FirstOrDefault(_ => _.CurrentFolder.FolderPath == folderPath);

            private Outlook.Explorer GetExplorer(Outlook.Folder folder)
                => folder.GetExplorer();
        }
    }
}
