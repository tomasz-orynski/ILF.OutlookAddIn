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
            }

            private static Logger _logger = LogManager.GetCurrentClassLogger();
            private readonly IReadOnlyDictionary<string, IFolderSource> _folders;
            private readonly Action<Action<ICW<Outlook.NavigationFolder>>> _iterator;

            public _FoldersSource(
                ICW<Outlook.Folder> rootFolder,
                Func<string, bool> folderFilter,
                Func<string, bool> folderSelected
                )
            {
                Contract.Assert(rootFolder != null);
                Contract.Assert(folderFilter != null);
                Contract.Assert(folderSelected != null);

                _iterator = action =>
                {
                    using (var explorer = GetExplorer(rootFolder))
                    using (var navPane = explorer.Call(_ => _.NavigationPane))
                    using (var mods = navPane.Call(_ => _.Modules))
                    using (var navMod = mods.Call(_ => _.GetNavigationModule(Outlook.OlNavigationModuleType.olModuleCalendar).As<Outlook.CalendarModule>()))
                    using (var navGrps = navMod.Call(_ => _.NavigationGroups))
                        navGrps.ForEach((ICW<Outlook.NavigationGroup> navGrp) =>
                        {
                            using (var navFlds = navGrp.Call(_ => _.NavigationFolders))
                                navFlds.ForEach(action);
                        });

                    //hack
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                };

                var fldSrc = new Dictionary<string, IFolderSource>();
                _folders = fldSrc;
                _iterator(navFld =>
                {
                    var name = navFld.Ref.DisplayName;
                    if (folderFilter(name))
                        using (var fld = navFld.Call(_ => _.Folder))
                        {
                            var id = fld.Ref.FolderPath;
                            if (!fldSrc.ContainsKey(id))
                                fldSrc.Add(id, new FolderSource()
                                {
                                    ID = id,
                                    Name = name,
                                    IsSelected = folderSelected(name),
                                });
                        }
                });
            }

            public IReadOnlyDictionary<string, IFolderSource> Folders => _folders;
            public void OnFolders(IReadOnlyDictionary<string, IFolderSource> folders, Action<IFolderSource, ICW<Outlook.MAPIFolder>> action)
            {
                _iterator(navFld =>
                {
                    var name = navFld.Ref.DisplayName;
                        using (var fld = navFld.Call(_ => _.Folder))
                        {
                            var id = fld.Ref.FolderPath;
                            if (folders.TryGetValue(id, out var fldSrc))
                                action(fldSrc, fld);
                        }
                });
            }

            private static ICW<Outlook.Explorer> GetExplorer(ICW<Outlook.Folder> folder)
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
