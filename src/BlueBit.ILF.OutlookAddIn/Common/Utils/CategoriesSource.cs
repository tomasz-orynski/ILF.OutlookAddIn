using BlueBit.ILF.OutlookAddIn.Common.Extensions;
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
        private readonly Outlook.Application _application;
        private readonly IEnumerable<Outlook.Category> _categoriesSource;
        private readonly IEnumerable<Action> _onDisposeActions;

        public CategoriesSource(
            Outlook.Categories categories
            )
        {
            Contract.Assert(categories != null);

            var onDisposeActions = new List<Action>();
            _onDisposeActions = onDisposeActions;

            _categoriesSource = categories
                .Cast<Outlook.Category>()
                .DebugFetch()
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
    }
}
