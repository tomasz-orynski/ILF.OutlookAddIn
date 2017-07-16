﻿using GalaSoft.MvvmLight;

namespace BlueBit.ILF.OutlookAddIn.MVVM.Models
{
    public class CategoryModel : ObservableObject
    {
        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set => Set(() => IsSelected, ref _isSelected, value);
        }

        public string ID { get; }
        public string Name { get;  }

        public CategoryModel((string id, string name) category)
        {
            ID = category.id;
            Name = category.name;
        }
    }
}
