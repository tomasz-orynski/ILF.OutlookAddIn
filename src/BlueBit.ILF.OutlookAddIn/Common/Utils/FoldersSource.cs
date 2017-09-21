using BlueBit.ILF.OutlookAddIn.Common.Extensions;
using BlueBit.ILF.OutlookAddIn.Diagnostics;
using MoreLinq;
using NLog;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.Contracts;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BlueBit.ILF.OutlookAddIn.Common.Utils
{
    class FoldersSource :
            IDisposable
    {
        private static Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly Outlook.Application _application;
        private readonly IEnumerable<(Outlook.Folder Folder, bool IsSelected)> _foldersSource;
        private readonly IEnumerable<Action> _onDisposeActions;

        public FoldersSource(
            Outlook.Folder rootFolder,
            Func<string, bool> folderFilter,
            Func<string, bool> folderSelected
            )
        {
            Contract.Assert(rootFolder != null);
            Contract.Assert(folderFilter != null);
            Contract.Assert(folderSelected != null);

            var onDisposeActions = new List<Action>();
            _onDisposeActions = onDisposeActions;
            _application = rootFolder.Application;

            _foldersSource = rootFolder.Application
                .GetNamespace("MAPI")
                .Stores
                .Cast<Outlook.Store>()
                .SelectMany(_ => _.GetRootFolder().Folders.Cast<Outlook.Folder>())
                .SafeWhere(_ => _.DefaultMessageClass == "IPM.Appointment")
                .SafeWhere(_ => folderFilter(_.Name))
                .OrderBy(_ => _.Name)
                .Select(_ => (_, folderSelected(_.Name)))
                .SafeToList()
                ;
        }

        public void Dispose()
        {
            _onDisposeActions.ForEach(_ => _.Invoke());
        }

        public void EnumFolders(Action<Outlook.Folder, bool> enumAction)
        {
            Contract.Assert(enumAction != null);
            _foldersSource
                .ForEach(_ => enumAction(_.Folder, _.IsSelected));
        }

        public IEnumerable<Outlook.Folder> GetFolders()
            => _foldersSource.Select(_ => _.Folder);
    }
}
