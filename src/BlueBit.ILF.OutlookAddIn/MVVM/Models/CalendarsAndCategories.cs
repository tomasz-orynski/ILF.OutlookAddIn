using BlueBit.ILF.OutlookAddIn.Properties;
using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using System;
using System.Collections.ObjectModel;
using System.Diagnostics.Contracts;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BlueBit.ILF.OutlookAddIn.MVVM.Models
{
    public class CategoryModel : ObservableObject
    {
        private readonly Outlook.Category _category;
        public Outlook.Category Category => _category;


        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set => Set(() => IsSelected, ref _isSelected, value);
        }

        public string Name => _category.Name;
        public string ID => _category.CategoryID;
        //public int Color => _category.Color;

        public CategoryModel(Outlook.Category category)
        {
            Contract.Assert(category != null);
            _category = category;
        }
    }


    public class CalendarsAndCategoriesModel
    {
        private readonly ObservableCollection<CalendarModel> _calendars = new ObservableCollection<CalendarModel>();
        private readonly ObservableCollection<CategoryModel> _categories = new ObservableCollection<CategoryModel>();
        private readonly ObservableCollection<ActionModel> _actions = new ObservableCollection<ActionModel>();

        public ObservableCollection<CalendarModel> Calendars => _calendars;
        public ObservableCollection<CategoryModel> Categories => _categories;
        public ObservableCollection<ActionModel> Actions => _actions;


        public CalendarsAndCategoriesModel(
            Action<Action<Outlook.Folder, bool>> foldersEnumerator,
            Action<Action<Outlook.Category>> categoriesEnumerator,
            Action<CalendarsAndCategoriesModel> cmdApply,
            Action<CalendarsAndCategoriesModel> cmdCancel)
        {
            Contract.Assert(foldersEnumerator != null);
            Contract.Assert(categoriesEnumerator != null);
            Contract.Assert(cmdApply != null);
            Contract.Assert(cmdCancel != null);


            _actions.Add(new ActionModel()
            {
                Command = new RelayCommand(() => cmdCancel(this)),
                Name = Resources.CmdCancel,
                IsCancel = true,
            });

            foldersEnumerator((folder, isSelected) => _calendars.Add(new CalendarModel(folder) { IsSelected = isSelected }));
            categoriesEnumerator(category => _categories.Add(new CategoryModel(category)));

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
