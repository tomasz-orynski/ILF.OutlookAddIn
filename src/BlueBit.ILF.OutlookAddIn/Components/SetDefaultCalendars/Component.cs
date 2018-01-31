using BlueBit.ILF.OutlookAddIn.Common.Extensions;
using BlueBit.ILF.OutlookAddIn.Common.Utils;
using BlueBit.ILF.OutlookAddIn.Diagnostics;
using BlueBit.ILF.OutlookAddIn.MVVM.Models;
using BlueBit.ILF.OutlookAddIn.MVVM.Views;
using BlueBit.ILF.OutlookAddIn.Properties;
using NLog;
using System;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BlueBit.ILF.OutlookAddIn.Components.SetDefaultCalendars
{
    public class Component : 
        IComponent,
        ICommandComponent
    {
        private static Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly IEnviroment _env;
        private readonly IConfiguration _cfg;

        public Component(
            IEnviroment env,
            IConfiguration cfg)
        {
            _env = env;
            _cfg = cfg;
        }

        public CommandID ID => CommandID.SetDefaultCalendars;

        public void Execute()
        {
            var window = new CalendarsWindow();
            window.Title = Resources.SetDefaultCalendars_Caption;
            window.DataContext = new CalendarsModel(
                _env,
                FuncExtensions
                    .IfTrueThenCloseWindow<CalendarsModel>(OnApply, window),
                FuncExtensions
                    .AlwaysTrue<CalendarsModel>()
                    .IfTrueThenCloseWindow(window)
                );
            window.ShowDialog(_logger);
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
