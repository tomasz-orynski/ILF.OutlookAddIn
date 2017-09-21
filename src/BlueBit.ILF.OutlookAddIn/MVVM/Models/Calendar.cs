﻿using BlueBit.ILF.OutlookAddIn.Common.Extensions.ForOutlook;
using GalaSoft.MvvmLight;
using System;
using System.Collections.ObjectModel;
using System.Diagnostics.Contracts;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BlueBit.ILF.OutlookAddIn.MVVM.Models
{
    public class CalendarModel : ObservableObject
    {
        public event Action<CalendarModel> SelectedChanged;


        private readonly Outlook.Folder _folder;
        public Outlook.Folder Folder => _folder;

        private readonly Lazy<ObservableCollection<CategoryModel>> _categories;
        public ObservableCollection<CategoryModel> Categories => _categories.Value;

        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set { if (Set(() => IsSelected, ref _isSelected, value)) SelectedChanged?.Invoke(this); }
        }

        public string Name => _folder.Name;

        public CalendarModel(Outlook.Folder folder)
        {
            Contract.Assert(folder != null);
            _folder = folder;
            _categories = new Lazy<ObservableCollection<CategoryModel>>(GetCategories);
        }

        private ObservableCollection<CategoryModel> GetCategories()
            => new ObservableCollection<CategoryModel>(
                _folder
                .GetCategories()
                .Select(_ => new CategoryModel(_)));
    }
}
