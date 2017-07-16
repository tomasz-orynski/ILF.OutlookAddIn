using BlueBit.ILF.OutlookAddIn.Diagnostics;
using BlueBit.ILF.OutlookAddIn.Properties;
using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.CommandWpf;
using NLog;
using System;
using System.Collections.ObjectModel;
using System.Diagnostics.Contracts;

namespace BlueBit.ILF.OutlookAddIn.MVVM.Models
{
    public abstract class RootModel<T> : ObservableObject
        where T: RootModel<T>
    {
        private static Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly ObservableCollection<ActionModel> _actions = new ObservableCollection<ActionModel>();

        public ObservableCollection<ActionModel> Actions => _actions;
        protected Action<T> CmdApply { get; set; }
        protected Action<T> CmdCancel { get; set; }

        protected abstract T This { get; }

        protected RootModel()
        {
            AddAction(Resources.CmdApply, OnCmdApply, IsCmdApply, isDefault: true);
            AddAction(Resources.CmdCancel, OnCmdCancel, IsCmdCancel, isCancel: true);
        }

        private void AddAction(string name, Action action, Func<bool> isAction, bool isDefault = false, bool isCancel = false)
        {
            Contract.Assert(!string.IsNullOrWhiteSpace(name));
            Contract.Assert(action != null);

            var model = new ActionModel()
            {
                Command = new RelayCommand(action, isAction),
                Name = name,
                IsDefault = isDefault,
                IsCancel = isCancel,
            };
            _actions.Add(model);
        }

        private bool IsCmdCancel()
            => CmdCancel != null;

        private void OnCmdCancel()
            => _logger.OnEntryCall(() => CmdCancel(This));

        private bool IsCmdApply()
            => CmdApply != null;

        private void OnCmdApply()
            => _logger.OnEntryCall(() => CmdApply(This));
    }
}
