using BlueBit.ILF.OutlookAddIn.Common.Extensions.ForOutlook;
using BlueBit.ILF.OutlookAddIn.Common.Patterns;
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

        private readonly string _folderPath;

        public string ID => _folderPath;
        public string Name { get; }
        public ObservableCollection<CategoryModel> Categories { get; }

        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set { if (Set(() => IsSelected, ref _isSelected, value)) SelectedChanged?.Invoke(this); }
        }

        public CalendarModel(ICW<Outlook.NavigationFolder> folder, IEnviroment env)
        {
            Contract.Assert(folder != null);
            Name = folder.Ref.DisplayName;
            using (var fld = folder.Call(_ => _.Folder))
            {
                _folderPath = fld.Ref.FolderPath;
                Categories = new ObservableCollection<CategoryModel>(env.GetCategories(fld).Select(_ => new CategoryModel(_)));
            }
        }

        public static string GetID(ICW<Outlook.NavigationFolder> folder)
        {
            using (var fld = folder.Call(_ => _.Folder))
                return fld.Ref.FolderPath;
        }
                
    }
}
