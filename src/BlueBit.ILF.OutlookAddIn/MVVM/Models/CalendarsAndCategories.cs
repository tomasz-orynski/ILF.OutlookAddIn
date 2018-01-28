using BlueBit.ILF.OutlookAddIn.Common.Extensions;
using BlueBit.ILF.OutlookAddIn.Components;
using BlueBit.ILF.OutlookAddIn.Properties;
using MoreLinq;
using System;
using System.Collections.ObjectModel;
using System.Diagnostics.Contracts;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BlueBit.ILF.OutlookAddIn.MVVM.Models
{
    public class CalendarsAndCategoriesModel :
        RootModel<CalendarsAndCategoriesModel>
    {
        protected override CalendarsAndCategoriesModel This => this;

        private readonly ObservableCollection<CalendarModel> _calendars = new ObservableCollection<CalendarModel>();
        public ObservableCollection<CalendarModel> Calendars => _calendars;

        private readonly ObservableCollection<CategoryModel> _categories = new ObservableCollection<CategoryModel>();
        public ObservableCollection<CategoryModel> Categories => _categories;

        public CalendarsAndCategoriesModel(
            IEnviroment env,
            Action<CalendarsAndCategoriesModel> cmdApply,
            Action<CalendarsAndCategoriesModel> cmdCancel)
        {
            Contract.Assert(env != null);
            Contract.Assert(cmdApply != null);
            Contract.Assert(cmdCancel != null);

            CmdCancel = cmdCancel;
            env.FoldersSource.EnumFolders((folder, isSelected) => _calendars.Add(new CalendarModel(folder, env) { IsSelected = isSelected }));

            if (_calendars.Count == 0)
                return;

            CmdApply = cmdApply;
            _calendars.ForEach(_ => {
                _.SelectedChanged += calendar =>
                {
                    if (calendar.IsSelected)
                        calendar.Categories.ForEach(_categories.Add);
                    else
                        calendar.Categories.ForEachFunc(_categories.Remove);
                };
                if (_.IsSelected)
                    _.Categories.ForEach(_categories.Add);
            });
        }
    }
}
