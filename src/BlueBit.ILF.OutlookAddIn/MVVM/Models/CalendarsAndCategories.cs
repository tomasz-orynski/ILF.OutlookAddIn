using BlueBit.ILF.OutlookAddIn.Common.Extensions;
using BlueBit.ILF.OutlookAddIn.Properties;
using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using MoreLinq;
using System;
using System.Collections.ObjectModel;
using System.Diagnostics.Contracts;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BlueBit.ILF.OutlookAddIn.MVVM.Models
{
    public class CalendarsAndCategoriesModel
    {
        private readonly ObservableCollection<CalendarModel> _calendars = new ObservableCollection<CalendarModel>();
        private readonly ObservableCollection<CategoryModel> _categories = new ObservableCollection<CategoryModel>();
        private readonly ObservableCollection<ActionModel> _actions = new ObservableCollection<ActionModel>();

        public ObservableCollection<CalendarModel> Calendars => _calendars;
        public ObservableCollection<CategoryModel> Categories => _categories;
        public ObservableCollection<ActionModel> Actions => _actions;

        public CalendarsAndCategoriesModel(
            Action<Action<Outlook.NavigationFolder, bool>> foldersEnumerator,
            Action<CalendarsAndCategoriesModel> cmdApply,
            Action<CalendarsAndCategoriesModel> cmdCancel)
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

            if (_calendars.Count == 0)
                return;

            _actions.Insert(0, new ActionModel()
            {
                Command = new RelayCommand(() => cmdApply(this)),
                Name = Resources.CmdApply,
                IsDefault = true,
            });

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
