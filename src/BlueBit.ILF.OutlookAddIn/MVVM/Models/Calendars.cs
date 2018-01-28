using BlueBit.ILF.OutlookAddIn.Components;
using BlueBit.ILF.OutlookAddIn.Properties;
using System;
using System.Collections.ObjectModel;
using System.Diagnostics.Contracts;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BlueBit.ILF.OutlookAddIn.MVVM.Models
{
    public class CalendarsModel : 
        RootModel<CalendarsModel>
    {
        protected override CalendarsModel This => this;

        private readonly ObservableCollection<CalendarModel> _calendars = new ObservableCollection<CalendarModel>();
        public ObservableCollection<CalendarModel> Calendars => _calendars;

        public CalendarsModel(
            IEnviroment env,
            Action<CalendarsModel> cmdApply,
            Action<CalendarsModel> cmdCancel)
        {
            Contract.Assert(env != null);
            Contract.Assert(cmdApply != null);
            Contract.Assert(cmdCancel != null);

            CmdCancel = cmdCancel;
            env.FoldersSource.EnumFolders((folder, isSelected) => _calendars.Add(new CalendarModel(folder, env) { IsSelected = isSelected }));
            if (_calendars.Count > 0)
                CmdApply = cmdApply;
        }
    }
}
