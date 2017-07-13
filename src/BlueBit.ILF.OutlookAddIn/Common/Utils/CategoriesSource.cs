using BlueBit.ILF.OutlookAddIn.Diagnostics;
using MoreLinq;
using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BlueBit.ILF.OutlookAddIn.Common.Utils
{
    class CategoriesSource :
            IDisposable
    {
        private readonly IEnumerable<Outlook.Category> _categoriesSource;
        private readonly IEnumerable<Action> _onDisposeActions;

        public CategoriesSource(
            FoldersSource foldersSource
            )
        {
            Contract.Assert(foldersSource != null);

            var onDisposeActions = new List<Action>();
            _onDisposeActions = onDisposeActions;

            _categoriesSource = foldersSource
                .GetFolders()
                .SelectMany(GetCategories)
                .SafeToList()
                .OrderBy(_=>_.Name);
        }

        public void Dispose()
        {
            _onDisposeActions.ForEach(_ => _.Invoke());
        }

        public void EnumCategories(Action<Outlook.Category> enumAction)
        {
            Contract.Assert(enumAction != null);
            _categoriesSource
                .ForEach(enumAction);
        }

        private IEnumerable<Outlook.Category> GetCategories(Outlook.Folder folder)
        {
            var storage = folder.GetStorage("http://schemas.microsoft.com/mapi/proptag/0x7C080102", Outlook.OlStorageIdentifierType.olIdentifyByMessageClass);
            var xml = storage.PropertyAccessor.GetProperty("PR_ROAMING_XMLSTREAM");
            return null;
        }
    }
}
