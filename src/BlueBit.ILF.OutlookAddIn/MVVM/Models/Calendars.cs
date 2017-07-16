using BlueBit.ILF.OutlookAddIn.Properties;
using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using System;
using System.Collections.ObjectModel;
using System.Diagnostics.Contracts;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BlueBit.ILF.OutlookAddIn.MVVM.Models
{
    public class CalendarsModel : 
        ObservableObject
    {

        private readonly ObservableCollection<CalendarModel> _calendars = new ObservableCollection<CalendarModel>();
        private readonly ObservableCollection<ActionModel> _actions = new ObservableCollection<ActionModel>();

        public ObservableCollection<CalendarModel> Calendars => _calendars;
        public ObservableCollection<ActionModel> Actions => _actions;


        public CalendarsModel(
            Action<Action<Outlook.NavigationFolder,bool>> foldersEnumerator,
            Action<CalendarsModel> cmdApply,
            Action<CalendarsModel> cmdCancel)
        {
            Contract.Assert(foldersEnumerator != null);
            Contract.Assert(cmdApply != null);
            Contract.Assert(cmdCancel != null);

            _actions.Add(new ActionModel()
            {
                Command = new RelayCommand(() => cmdCancel(this)),
                Name = Resources.CmdCancel,
                IsCancel = true,
            });

            foldersEnumerator((folder, isSelected) => _calendars.Add(new CalendarModel(folder) { IsSelected = isSelected }));

            if (_calendars.Count > 0)
                _actions.Insert(0, new ActionModel()
                {
                    Command = new RelayCommand(() => cmdApply(this)),
                    Name = Resources.CmdApply,
                    IsDefault = true,
                });
        }
    }
}
