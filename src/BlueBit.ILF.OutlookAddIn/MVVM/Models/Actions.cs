using GalaSoft.MvvmLight;
using System.Windows.Input;

namespace BlueBit.ILF.OutlookAddIn.MVVM.Models
{
    public class ActionModel : ObservableObject
    {
        private bool _isDefault;
        public bool IsDefault
        {
            get => _isDefault;
            set => Set(() => IsDefault, ref _isDefault, value);
        }

        private bool _isCancel;
        public bool IsCancel
        {
            get => _isCancel;
            set => Set(() => IsCancel, ref _isCancel, value);
        }

        private string _name;
        public string Name
        {
            get => _name;
            set => Set(() => Name, ref _name, value);
        }

        private ICommand _command;
        public ICommand Command
        {
            get => _command;
            set => Set(() => Command, ref _command, value);
        }
    }
}
