using BlueBit.ILF.OutlookAddIn.Common.Extensions;
using BlueBit.ILF.OutlookAddIn.Common.Patterns;
using BlueBit.ILF.OutlookAddIn.Common.Utils;
using BlueBit.ILF.OutlookAddIn.Diagnostics;
using BlueBit.ILF.OutlookAddIn.Common.Extensions.ForOutlook;
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
            private class FolderSource : IFolderSource
            {
                public string ID { get; set; }
                public string Name { get; set; }
                public bool IsSelected { get; set; }
                public IReadOnlyList<(string ID, string Name)> Categories { get; set; }
            }

            private static Logger _logger = LogManager.GetCurrentClassLogger();
            private readonly ICW<Outlook.Folder> _rootFolder;
            private readonly IReadOnlyList<FolderSource> _folders;

            public _FoldersSource(
                ICW<Outlook.Folder> rootFolder,
                Func<string, bool> folderFilter,
                Func<string, bool> folderSelected
                )
            {
                Contract.Assert(rootFolder != null);
                Contract.Assert(folderFilter != null);
                Contract.Assert(folderSelected != null);

                _rootFolder = rootFolder;
                using (var explorer = GetExplorer(_rootFolder))
                using (var navPane = explorer.Call(_ => _.NavigationPane))
                using (var mods = navPane.Call(_ => _.Modules))
                using (var navMod = mods.Call(_ => _.GetNavigationModule(Outlook.OlNavigationModuleType.olModuleCalendar).As<Outlook.CalendarModule>()))
                using (var navGrps = navMod.Call(_ => _.NavigationGroups))
                {
                    var fldSrc = new List<FolderSource>();
                    _folders = fldSrc;

                    navGrps.ForEach((ICW<Outlook.NavigationGroup> navGrp) => {
                        using (var navFlds = navGrp.Call(_ => _.NavigationFolders))
                            navFlds.ForEach((ICW<Outlook.NavigationFolder> navFld) =>
                            {
                                var name = navFld.Ref.DisplayName;
                                if (folderFilter(name))
                                    using (var fld = navFld.Call(_ => _.Folder))
                                        fldSrc.Add(new FolderSource() {
                                            ID = fld.Ref.FolderPath,
                                            Name = name,
                                            IsSelected = folderSelected(name),
                                            Categories = fld.GetCategoriesFromTable().NullAsEmpty().ToList(),
                                        });
                            });
                    });
                }

                var tmp = GetFolder(_folders[0]);
            }

            public IReadOnlyList<IFolderSource> Folders => _folders;

            public ICW<Outlook.Folder> GetFolder(IFolderSource folderSource)
            {
                using (var rootItems = _rootFolder.Call(_ => _.Items))
                using (var rootItem = rootItems.Call(_ => (object)_[1]))
                {

                }
                return null;
            }

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
