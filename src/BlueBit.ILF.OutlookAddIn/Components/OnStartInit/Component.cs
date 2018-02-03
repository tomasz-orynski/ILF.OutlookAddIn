using BlueBit.ILF.OutlookAddIn.Common.Extensions;
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
        IInitializedAppComponent,
        IEnviroment
    {
        private static Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly IConfiguration _cfg;
        private string _userName;
        private ICW<Outlook.Application> _app;
        private ICW<Outlook.Folder> _calendarFolder;
        private ICW<Outlook.Items> _calendarItems;
        private IFoldersSource _foldersSource;

        public string UserName => _userName;
        public ICW<Outlook.Application> Application => _app;
        public ICW<Outlook.Folder> CalendarFolder => _calendarFolder;
        public ICW<Outlook.Items> CalendarItems => _calendarItems;
        public IFoldersSource FoldersSource => _foldersSource;

        public Component(IConfiguration cfg)
        {
            _cfg = cfg;
        }

        public void Initialize(Outlook.Application app)
            => _logger.OnEntryCall(() =>
        {
            _app = app.AsCW();
            using (var session = _app.Call(_ => _.Session))
            using (var user = session.Call(_ => _.CurrentUser))
            using (var ns = _app.Call(_ => _.GetNamespace("MAPI")))
            {
                _userName = user.Ref.Name;
                _calendarFolder = ns.Call(_ => (Outlook.Folder)_.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar));
                _calendarItems = _calendarFolder.Call(_ => _.Items);
            }

            GetHandlers()
                .Select(_ =>
                {
                    var timer = new DispatcherTimer();
                    timer.Interval = new TimeSpan(0, 0, _.Period);
                    return new { Timer = timer, _.Handler };
                })
                .AsNodes()
                .Select(_ =>
                {
                    _.Curr.Timer.Tick += (s, e) => _logger.OnEntryCall(() =>
                    {
                        _.Curr.Timer.Stop();
                        _.Curr.Handler();
                        _.Next?.Timer?.Start();
                    });
                    return _.Curr.Timer;
                })
                .ToList()
                .First()
                .Start();
        });

        IEnumerable<(int Period, Action Handler)> GetHandlers()
        {
            var initOnStart = _cfg.GetInitOnStart();
            if (initOnStart <= 0)
                yield break;

            yield return (
                initOnStart,
                () => SwitchToView(_calendarNames)
            );
            yield return (
                1,
                () => {
                    _foldersSource = OnCreateFoldersSource();
                }
            );
        }

        static readonly IReadOnlyList<string> _calendarNames = new List<string>() { "Kalendarz", "Calendar" };
        private void SwitchToView(IReadOnlyList<string> names)
        {
            using (var explorer = _app.Call(_ => _.ActiveExplorer()))
            using (var views = _calendarFolder.Call(_ => _.Views))
            {
                explorer.Ref.CurrentFolder = _calendarFolder.Ref;
                views.ForEach((ICW<Outlook.View> view) =>
                {
                    if (names.Contains(view.Ref.Name))
                    {
                        explorer.Ref.CurrentView = view.Ref;
                        view.Ref.Apply();
                    }
                });
            }
        }

        private _FoldersSource OnCreateFoldersSource()
            => _logger.OnEntryCall(() => new _FoldersSource(
                _calendarFolder,
                _cfg.GetCalendarPrefixes().AsPrefixFilter(),
                _cfg.GetDeafultCalendars().AsEqualsFilter()
                ));
    }
}
