using BlueBit.ILF.OutlookAddIn.Common.Extensions;
using BlueBit.ILF.OutlookAddIn.Diagnostics;
using BlueBit.ILF.OutlookAddIn.MVVM.Models;
using BlueBit.ILF.OutlookAddIn.MVVM.Views;
using MoreLinq;
using NLog;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BlueBit.ILF.OutlookAddIn.Components.OnAddAppointmentHandler
{
    public class Component : ISelfRegisteredComponent
    {
        private static Logger _logger = LogManager.GetCurrentClassLogger();
        private IConfiguration _cfg;

        public Component(IConfiguration cfg)
        {
            _cfg = cfg;
        }

        public void Initialize(Outlook.Application app)
        {
            app
                .GetNamespace("MAPI")
                .GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar)
                .Items
                .ItemAdd += HandlerExtensions.AsSingleItemAddEventHandler(items_ItemAdd);
        }

        private void items_ItemAdd(object item)
            => _logger.OnEntryCall(() =>
            {
                var appointment = item as Outlook.AppointmentItem;
                if (appointment == null) return;

                if (appointment.ResponseStatus != Outlook.OlResponseStatus.olResponseAccepted) return;
                var rootFolder = (Outlook.Folder)appointment.Parent;
                if (!rootFolder.FolderPath.StartsWith(@"\\")) return;

                var prefixes = _cfg.GetCalendarPrefixes().ToList();
                var selected = _cfg.GetDeafultCalendars().ToList();
                using (var foldersSource = new FoldersSource(
                    rootFolder,
                    name => prefixes.Any(name.StartsWith),
                    name => selected.Any(name.Equals)
                    ))
                {
                    var window = new CalendarsWindow();
                    window.DataContext = new CalendarsModel(
                        foldersSource.EnumFolders,
                        model => {
                            model.Calendars
                                .Where(_ => _.IsSelected)
                                .ForEach(_ =>
                                {
                                    Clone(_.Folder, appointment);
                                });
                            window.Close();
                        },
                        model => {
                            window.Close();
                        });
                    window.ShowDialog();
                }
            });

        private static Outlook.AppointmentItem Clone(Outlook.Folder folder, Outlook.AppointmentItem source)
        {
            var item = (Outlook.AppointmentItem)folder.Items.Add(Outlook.OlItemType.olAppointmentItem);
            item.Delete();
            item = source.Copy();
            item = item.Move(folder);
            item.Subject = source.Application.Session.CurrentUser.Name;
            item.Save();
            return item;
        }
    }
}
