using BlueBit.ILF.OutlookAddIn.Common.Extensions;
using BlueBit.ILF.OutlookAddIn.Common.Extensions.ForOutlook;
using BlueBit.ILF.OutlookAddIn.Common.Utils;
using BlueBit.ILF.OutlookAddIn.Components.OnAddAppointmentHandler;
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
            var getRootFolder = new Lazy<Outlook.Folder>(() => app
                .GetNamespace("MAPI")
                .GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar)
                .As<Outlook.Folder>());

            _categories = new Lazy<IReadOnlyDictionary<string, IReadOnlyList<(string id, string name)>>>(OnGetCategories);
            _foldersSource = new Lazy<IFoldersSource>(() => OnCreateFoldersSource(getRootFolder));

            if (_cfg.GetInitOnStart())
            {
                var timer = new DispatcherTimer();
                timer.Interval = new TimeSpan(0, 0, 15);
                timer.Tick += (s, e) =>
                {
                    timer.Stop();
                    OnTimer();
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
                    var folder = fld.Folder;
                    dict.Add(folder.FolderPath, folder.GetCategories().ToList());
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
