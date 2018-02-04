using BlueBit.ILF.OutlookAddIn.Components;
using System;
using System.Collections.ObjectModel;
using System.Diagnostics.Contracts;
using System.Linq;

namespace BlueBit.ILF.OutlookAddIn.MVVM.Models
{
    public class CalendarsModel : 
        RootModel<CalendarsModel>
    {
        protected override CalendarsModel This => this;

        private readonly Lazy<ObservableCollection<CalendarModel>> _calendars;
        public ObservableCollection<CalendarModel> Calendars => _calendars.Value;

        public CalendarsModel(
            IEnviroment env,
            Action<CalendarsModel> cmdApply,
            Action<CalendarsModel> cmdCancel)
        {
            Contract.Assert(env != null);
            Contract.Assert(cmdApply != null);
            Contract.Assert(cmdCancel != null);

            CmdCancel = cmdCancel;
            CmdApply = cmdApply;
            _calendars = new Lazy<ObservableCollection<CalendarModel>>(() => new ObservableCollection<CalendarModel>(env.FoldersSource.Folders.Select(_ => new CalendarModel(_.Value, env))));
        }
    }
}
