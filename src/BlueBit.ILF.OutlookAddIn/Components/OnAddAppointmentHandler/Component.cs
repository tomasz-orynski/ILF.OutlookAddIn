using BlueBit.ILF.OutlookAddIn.Common.Extensions;
using BlueBit.ILF.OutlookAddIn.Common.Patterns;
using BlueBit.ILF.OutlookAddIn.Diagnostics;
using BlueBit.ILF.OutlookAddIn.MVVM.Models;
using BlueBit.ILF.OutlookAddIn.MVVM.Views;
using BlueBit.ILF.OutlookAddIn.Properties;
using MoreLinq;
using NLog;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BlueBit.ILF.OutlookAddIn.Components.OnAddAppointmentHandler
{
    public class Component : IInitializedComponent
    {
        private static Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly IEnviroment _env;

        public Component(IEnviroment env)
        {
            _env = env;
        }

        public void Initialize()
        {
            _env.CalendarItems.Ref.ItemAdd += OnItemAdd;
        }

        private void OnItemAdd(object item)
            => _logger.OnSingleEntryCall<Component>(() =>
            {
                var appointment = item as Outlook.AppointmentItem;
                if (appointment == null) return;
                if (appointment.SafeCheck(_ => _.ResponseStatus != Outlook.OlResponseStatus.olResponseAccepted, true)) return;

                using (var rootFolder = appointment.Parent.AsCW_<Outlook.Folder>())
                    if (!rootFolder.Ref.FolderPath.StartsWith(@"\\")) return;

                var window = new CalendarsAndCategoriesWindow();
                window.DataContext = new CalendarsAndCategoriesModel(
                    _env,
                    FuncExtensions
                        .ApplyParams<CalendarsAndCategoriesModel, Outlook.AppointmentItem>(OnApply, appointment)
                        .IfTrueThenCloseWindow(window),
                    FuncExtensions
                        .AlwaysTrue<CalendarsAndCategoriesModel>()
                        .IfTrueThenCloseWindow(window)
                    );
                window.Title = Resources.OnAddAppointmentHandler_Caption;
                window.ShowDialog(_logger);
            });

        private bool OnApply(CalendarsAndCategoriesModel model, Outlook.AppointmentItem appointment)
        {
            var dict = model.Calendars
                .Where(_ => _.IsSelected)
                .ToDictionary(_ => _.ID, c => string.Join(",", c.Categories.Where(_ => _.IsSelected).OrderBy(_ => _.Name).Select(_ => _.Name)));
            _env.FoldersSource.EnumFolders((folder, sel) =>
            {
                var id = CalendarModel.GetID(folder);
                if (dict.TryGetValue(id, out var categories))
                {
                    using (var fld = folder.Call(_ => _.Folder))
                        Clone(fld, appointment, categories);
                }
            });
            return true;
        }

        private void Clone(
            ICW<Outlook.MAPIFolder> folder, 
            Outlook.AppointmentItem source, 
            string categories)
        {
            using (var item = source.CopyTo(folder.Ref, Outlook.OlAppointmentCopyOptions.olCreateAppointment).AsCW())
            {
                item.Ref.Subject = _env.UserName;
                item.Ref.Categories = categories;
                item.Ref.Save();
            }
        }
    }
}
