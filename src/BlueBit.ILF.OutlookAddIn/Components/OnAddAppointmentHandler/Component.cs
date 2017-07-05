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
            => _logger.OnEntryCall(HandlerExtensions.AsSingleEventHandler(() =>
            {
                var appointment = item as Outlook.AppointmentItem;
                if (appointment == null) return;

                if (appointment.ResponseStatus != Outlook.OlResponseStatus.olResponseAccepted) return;
                var rootFolder = (Outlook.Folder)appointment.Parent;
                if (!rootFolder.FolderPath.StartsWith(@"\\")) return;
                var categories = appointment.Application.Session.Categories;

                using (var foldersSource = new FoldersSource(
                    rootFolder,
                    _cfg.GetCalendarPrefixes().AsPrefixFilter(),
                    _cfg.GetDeafultCalendars().AsEqualsFilter()
                    ))
                using (var categoriesSource = new CategoriesSource(categories))
                {
                    var window = new CalendarsAndCategoriesWindow();
                    window.DataContext = new CalendarsAndCategoriesModel(
                        foldersSource.EnumFolders,
                        categoriesSource.EnumCategories,
                        FuncExtensions
                            .ApplyParams<CalendarsAndCategoriesModel, Outlook.AppointmentItem>(OnApply, appointment)
                            .IfTrueThenCloseWindow(window),
                        FuncExtensions
                            .AlwaysTrue<CalendarsAndCategoriesModel>()
                            .IfTrueThenCloseWindow(window)
                        );
                    window.Title = Resources.OnAddAppointmentHandler_Caption;
                    window.ShowDialog();
                }
            }));

        private bool OnApply(CalendarsAndCategoriesModel model, Outlook.AppointmentItem appointment)
        {
            var categories = string.Join(",", model.Categories.Select(_ => _.ID));
            model.Calendars
                .Where(_ => _.IsSelected)
                .ForEach(_ => Clone(_.Folder.Folder.As<Outlook.Folder>(), appointment, categories));
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
