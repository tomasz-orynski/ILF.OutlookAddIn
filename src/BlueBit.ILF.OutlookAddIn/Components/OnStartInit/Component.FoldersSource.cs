using BlueBit.ILF.OutlookAddIn.Common.Extensions;
using BlueBit.ILF.OutlookAddIn.Common.Patterns;
using BlueBit.ILF.OutlookAddIn.Common.Utils;
using BlueBit.ILF.OutlookAddIn.Diagnostics;
using MoreLinq;
using NLog;
using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BlueBit.ILF.OutlookAddIn.Components.OnStartInit
{
    partial class Component
    {
        private class _FoldersSource : IFoldersSource
        {
            private static Logger _logger = LogManager.GetCurrentClassLogger();
            private readonly IEnumerable<Tuple<Outlook.NavigationFolder, bool>> _foldersSource;

            public _FoldersSource(
                ICW<Outlook.Folder> rootFolder,
                Func<string, bool> folderFilter,
                Func<string, bool> folderSelected
                )
            {
                Contract.Assert(rootFolder != null);
                Contract.Assert(folderFilter != null);
                Contract.Assert(folderSelected != null);

                using (var explorer = GetExplorer(rootFolder))
                using (var navPane = explorer.Call(_ => _.NavigationPane))
                using (var mods = navPane.Call(_ => _.Modules))
                using (var navMod = mods.Call(_ => _.GetNavigationModule(Outlook.OlNavigationModuleType.olModuleCalendar).As<Outlook.CalendarModule>()))
                using (var navGrps = navMod.Call(_ => _.NavigationGroups))
                {
                    /* TODO-TO
                    _foldersSource = (GetExplorer(rootFolder.FolderPath) ?? rootFolder.Call(_ => _.GetExplorer()))
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
                    */
                }
            }

            public void EnumFolders(Action<Outlook.NavigationFolder, bool> enumAction)
                => _logger.OnEntryCall(() =>
                {
                    Contract.Assert(enumAction != null);
                    _foldersSource
                        .ForEach(_ => enumAction(_.Item1, _.Item2));
                });

            private ICW<Outlook.Explorer> GetExplorer(ICW<Outlook.Folder> folder)
            {
                using (var app = folder.Call(_ => _.Application))
                using (var explorers = app.Call(_ => _.Explorers))
                {
                    foreach (Outlook.Explorer explorer in explorers.Ref)
                    {
                        using (var expFld = explorer.CurrentFolder.AsCW())
                            if (expFld.Ref.FolderPath == folder.Ref.FolderPath)
                                return explorer.AsCW();
                        Marshal.ReleaseComObject(explorer);
                    }
                }
                return folder.Call(_ => _.GetExplorer());
            }
        }
    }
}
