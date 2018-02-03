using BlueBit.ILF.OutlookAddIn.Common.Patterns;
using BlueBit.ILF.OutlookAddIn.Common.Utils;
using BlueBit.ILF.OutlookAddIn.Components;
using GalaSoft.MvvmLight;
using System;
using System.Collections.ObjectModel;
using System.Diagnostics.Contracts;
using System.Linq;

namespace BlueBit.ILF.OutlookAddIn.MVVM.Models
{
    public class CalendarModel : ObservableObject
    {
        public event Action<CalendarModel> SelectedChanged;

        private readonly IFolderSource _folder;
        private readonly Lazy<ObservableCollection<CategoryModel>> _categories;

        public IFolderSource Folder => _folder;
        public string Name => _folder.Name;
        public ObservableCollection<CategoryModel> Categories => _categories.Value;

        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set { if (Set(() => IsSelected, ref _isSelected, value)) SelectedChanged?.Invoke(this); }
        }

        public CalendarModel(IFolderSource folder, IEnviroment env)
        {
            Contract.Assert(folder != null);
            _folder = folder;
            _categories = new Lazy<ObservableCollection<CategoryModel>>(() => new ObservableCollection<CategoryModel>(_folder.Categories.Select(_ => new CategoryModel(_))));
            _isSelected = folder.IsSelected;
        }
    }
}
