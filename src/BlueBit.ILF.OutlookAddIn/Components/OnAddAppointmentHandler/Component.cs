using BlueBit.ILF.OutlookAddIn.Common.Extensions;
using BlueBit.ILF.OutlookAddIn.Common.Utils;
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
    public class Component : ISelfRegisteredComponent
    {
        private static Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly IConfiguration _cfg;
        private Outlook.Items _items;

        public Component(IConfiguration cfg)
        {
            _cfg = cfg;
        }

        public void Initialize(Outlook.Application app)
        {
            _items = app
                .GetNamespace("MAPI")
                .GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar)
                .Items;
            _items.ItemAdd += OnItemAdd;
        }

        private void OnItemAdd(object item)
            => _logger.OnSingleEntryCall<Component>(() =>
            {
                var appointment = item as Outlook.AppointmentItem;
                if (appointment == null) return;

                if (appointment.SafeCheck(_ => _.ResponseStatus != Outlook.OlResponseStatus.olResponseAccepted, true)) return;
                var rootFolder = (Outlook.Folder)appointment.Parent;
                if (!rootFolder.FolderPath.StartsWith(@"\\")) return;

                using (var foldersSource = new FoldersSource(
                    rootFolder,
                    _cfg.GetCalendarPrefixes().AsPrefixFilter(),
                    _cfg.GetDeafultCalendars().AsEqualsFilter()
                    ))
                {
                    var window = new CalendarsAndCategoriesWindow();
                    window.DataContext = new CalendarsAndCategoriesModel(
                        foldersSource.EnumFolders,
                        FuncExtensions
                            .ApplyParams<CalendarsAndCategoriesModel, Outlook.AppointmentItem>(OnApply, appointment)
                            .IfTrueThenCloseWindow(window),
                        FuncExtensions
                            .AlwaysTrue<CalendarsAndCategoriesModel>()
                            .IfTrueThenCloseWindow(window)
                        );
                    window.Title = Resources.OnAddAppointmentHandler_Caption;
                    window.ShowDialog(_logger);
                }
            });

        private bool OnApply(CalendarsAndCategoriesModel model, Outlook.AppointmentItem appointment)
        {
            model.Calendars
                .Where(_ => _.IsSelected)
                .ForEach(c => Clone(
                    c.Folder.Folder.As<Outlook.Folder>(), 
                    appointment, 
                    string.Join(",", c.Categories.Where(_ => _.IsSelected).OrderBy(_ => _.Name).Select(_ => _.Name))));
            return true;
        }

        private static Outlook.AppointmentItem Clone(Outlook.Folder folder, Outlook.AppointmentItem source, string categories)
        {
            var item = (Outlook.AppointmentItem)folder.Items.Add(Outlook.OlItemType.olAppointmentItem);
            item.Delete();
            item = (Outlook.AppointmentItem)source.Copy();
            item = (Outlook.AppointmentItem)item.Move(folder);
            item.Subject = source.Application.Session.CurrentUser.Name;
            item.Categories = categories;
            item.Save();
            return item;
        }
    }
}
