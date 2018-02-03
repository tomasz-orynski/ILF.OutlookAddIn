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
                private ICW<Outlook.MAPIFolder> _folder;
                private Lazy<IReadOnlyList<(string ID, string Name)>> _categories;

                public ICW<Outlook.MAPIFolder> Folder => _folder;
                public string ID => _folder.Ref.FolderPath;
                public string Name { get; set; }
                public bool IsSelected { get; set; }

                public IReadOnlyList<(string ID, string Name)> Categories => _categories.Value;

                public FolderSource(ICW<Outlook.MAPIFolder> folder)
                {
                    _folder = folder;
                    _categories = new Lazy<IReadOnlyList<(string ID, string Name)>>(() => _folder.GetCategoriesFromTable().NullAsEmpty().ToList());
                }
            }

            private static Logger _logger = LogManager.GetCurrentClassLogger();
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

                using (var explorer = GetExplorer(rootFolder))
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
                                        fldSrc.Add(new FolderSource(navFld.Call(_ => _.Folder)) {
                                            Name = name,
                                            IsSelected = folderSelected(name),
                                        });
                            });
                    });
                }
            }

            public IReadOnlyList<IFolderSource> Folders => _folders;

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
