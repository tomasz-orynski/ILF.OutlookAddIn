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

                var fldSrc = new List<FolderSource>();
                rootFolder
                    .Call(_ => _.Application)
                    .Call_(_ => _.Explorers)
                    .ForEach_((ICW<Outlook.Explorer> explorer) =>
                    {
                        explorer
                            .Call(_ => _.NavigationPane)
                            .Call_(_ => _.Modules)
                            .Call_(_ => (Outlook.CalendarModule)_.GetNavigationModule(Outlook.OlNavigationModuleType.olModuleCalendar))
                            .Call_(_ => _.NavigationGroups)
                            .ForEach_((ICW<Outlook.NavigationGroup> navGrp) =>
                            {
                                navGrp
                                    .Call(_ => _.NavigationFolders)
                                    .ForEach_((ICW<Outlook.NavigationFolder> navFld) =>
                                    {
                                        var name = navFld.Ref.DisplayName;
                                        if (folderFilter(name))
                                            fldSrc.Add(new FolderSource(navFld.Call(_ => _.Folder))
                                            {
                                                Name = name,
                                                IsSelected = folderSelected(name),
                                            });
                                    });
                            });
                    });
                _folders = fldSrc;
            }

            public IReadOnlyList<IFolderSource> Folders => _folders;
        }
    }
}
