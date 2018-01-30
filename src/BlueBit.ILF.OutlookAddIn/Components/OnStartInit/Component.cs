using BlueBit.ILF.OutlookAddIn.Common.Extensions;
using BlueBit.ILF.OutlookAddIn.Common.Extensions.ForOutlook;
using BlueBit.ILF.OutlookAddIn.Common.Patterns;
using BlueBit.ILF.OutlookAddIn.Common.Utils;
using BlueBit.ILF.OutlookAddIn.Diagnostics;
using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BlueBit.ILF.OutlookAddIn.Components.OnStartInit
{
    public partial class Component :
        ISelfRegisteredComponent,
        IEnviroment
    {
        private static Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly IConfiguration _cfg;
        private Lazy<IReadOnlyDictionary<string, IReadOnlyList<(string id, string name)>>> _categories;
        private Lazy<IFoldersSource> _foldersSource;

        public IReadOnlyList<(string id, string name)> GetCategories(Outlook.MAPIFolder folder) => _categories.Value[folder.FolderPath];
        public IFoldersSource FoldersSource => _foldersSource.Value;


        public Component(IConfiguration cfg)
        {
            _cfg = cfg;
        }

        public void Initialize(Outlook.Application app)
        {
            var names = new List<string>() { "Calendar", "Kalendarz" };
            var getRootFolder = new Lazy<ICW<Outlook.Folder>>(() =>
            {
                var fld = app
                    .GetNamespace("MAPI")
                    .GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                var explorer = app.ActiveExplorer();
                explorer.CurrentFolder = fld;
                foreach (Outlook.View view in fld.Views)
                {
                    if (names.Contains(view.Name))
                    {
                        var calView = (Outlook.CalendarView)view;
                        explorer.CurrentView = calView;
                        calView.Apply();
                        break;
                    }
                }
                return fld.As<Outlook.Folder>();
            });

            _categories = new Lazy<IReadOnlyDictionary<string, IReadOnlyList<(string id, string name)>>>(OnGetCategories);
            _foldersSource = new Lazy<IFoldersSource>(() => OnCreateFoldersSource(getRootFolder));

            var initOnStart = _cfg.GetInitOnStart();
            if (initOnStart > 0)
            {
                var timer = new DispatcherTimer();
                timer.Interval = new TimeSpan(0, 0, initOnStart);
                var timer2 = new DispatcherTimer();
                timer2.Interval = new TimeSpan(0, 0, 1);
                timer2.Tick += (s, e) =>
                {
                    timer2.Stop();
                    OnTimer();
                };
                timer.Tick += (s, e) =>
                {
                    timer.Stop();
                    var fld = getRootFolder.Value;
                    timer2.Start();
                };
                timer.Start();
            }
        }

        public void Execute()
        {
        }

        private Dictionary<string, IReadOnlyList<(string id, string name)>> OnGetCategories()
            => _logger.OnEntryCall(() =>
            {
                var dict = new Dictionary<string, IReadOnlyList<(string id, string name)>>();
                _foldersSource.Value.EnumFolders((fld, sel) =>
                {
                    using (var folder = fld.Folder.AsCW())
                        dict[folder.Ref.FolderPath] = folder.GetCategoriesFromTable().NullAsEmpty().ToList();
                });
                return dict;
            });

        private _FoldersSource OnCreateFoldersSource(Lazy<Outlook.Folder> folder)
            => _logger.OnEntryCall(() => new _FoldersSource(folder.Value,
                _cfg.GetCalendarPrefixes().AsPrefixFilter(),
                _cfg.GetDeafultCalendars().AsEqualsFilter()
                ));

        private void OnTimer()
            => _logger.OnEntryCall(() => {
                if (!_categories.IsValueCreated)
                {
                    var categories = _categories.Value;
                }
            });
    }
}
