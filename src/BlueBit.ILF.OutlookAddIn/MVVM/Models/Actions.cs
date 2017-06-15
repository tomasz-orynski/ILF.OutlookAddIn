using GalaSoft.MvvmLight;
using System.Windows.Input;

namespace BlueBit.ILF.OutlookAddIn.MVVM.Models
{
    public class ActionModel : ObservableObject
    {
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
