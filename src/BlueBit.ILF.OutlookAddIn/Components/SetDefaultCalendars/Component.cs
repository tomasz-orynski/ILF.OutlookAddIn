using BlueBit.ILF.OutlookAddIn.Common.Extensions;
using BlueBit.ILF.OutlookAddIn.Common.Utils;
using BlueBit.ILF.OutlookAddIn.MVVM.Models;
using BlueBit.ILF.OutlookAddIn.MVVM.Views;
using BlueBit.ILF.OutlookAddIn.Properties;
using System;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BlueBit.ILF.OutlookAddIn.Components.SetDefaultCalendars
{
    public class Component : 
        ISelfRegisteredComponent,
        ICommandComponent
    {
        private readonly IConfiguration _cfg;
        private Func<Outlook.Folder> _getRootFolder;

        public Component(IConfiguration cfg)
        {
            _cfg = cfg;
        }

        public void Initialize(Outlook.Application app)
        {
            _getRootFolder = app
                .GetNamespace("MAPI")
                .GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar)
                .As<Outlook.Folder>;
        }

        public CommandID ID => CommandID.SetDefaultCalendars;

        public void Execute()
        {
            using (var foldersSource = new FoldersSource(
                _getRootFolder(),
                _cfg.GetCalendarPrefixes().AsPrefixFilter(),
                _cfg.GetDeafultCalendars().AsEqualsFilter()
                ))
            {
                var window = new CalendarsWindow();
                window.Title = Resources.SetDefaultCalendars_Caption;
                window.DataContext = new CalendarsModel(
                    foldersSource.EnumFolders,
                    FuncExtensions
                        .IfTrueThenCloseWindow<CalendarsModel>(OnApply, window),
                    FuncExtensions
                        .AlwaysTrue<CalendarsModel>()
                        .IfTrueThenCloseWindow(window)
                    );
                window.ShowDialog();
            }
        }

        private bool OnApply(CalendarsModel model)
        {
            _cfg.SetDeafultCalendars(
                model.Calendars
                    .Where(_ => _.IsSelected)
                    .Select(_ => _.Name));
            return true;
        }
    }
}
