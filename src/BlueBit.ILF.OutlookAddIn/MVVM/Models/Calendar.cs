﻿using BlueBit.ILF.OutlookAddIn.Common.Extensions.ForOutlook;
using BlueBit.ILF.OutlookAddIn.Components;
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

        private readonly Outlook.NavigationFolder _folder;
        public Outlook.NavigationFolder Folder => _folder;

        private readonly ObservableCollection<CategoryModel> _categories;
        public ObservableCollection<CategoryModel> Categories => _categories;

        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set { if (Set(() => IsSelected, ref _isSelected, value)) SelectedChanged?.Invoke(this); }
        }

        public string Name => _folder.DisplayName;

        public CalendarModel(Outlook.NavigationFolder folder, IEnviroment env)
        {
            Contract.Assert(folder != null);
            _folder = folder;
            _categories = new ObservableCollection<CategoryModel>(env.GetCategories(folder.Folder).Select(_ => new CategoryModel(_)));
        }
    }
}
