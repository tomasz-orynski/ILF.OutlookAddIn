using BlueBit.ILF.OutlookAddIn.Diagnostics;
using MoreLinq;
using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BlueBit.ILF.OutlookAddIn.Common.Utils
{
    class CategoriesSource :
            IDisposable
    {
        private readonly IEnumerable<Outlook.Category> _categoriesSource;

        public CategoriesSource(
            FoldersSource foldersSource
            )
        {
            Contract.Assert(foldersSource != null);

            var folders = foldersSource
                .GetFolders()
                .SafeToList();

            _categoriesSource = folders
                .SelectMany(GetCategories)
                .SafeToList()
                .OrderBy(_=>_.Name);
        }

        public void Dispose()
        {
        }

        public void EnumCategories(Action<Outlook.Category> enumAction)
        {
            Contract.Assert(enumAction != null);
            _categoriesSource
                .ForEach(enumAction);
        }


        static class Columns
        {
            public const string Property = "http://schemas.microsoft.com/mapi/proptag/0x7C080102";
            public const string MessageClass = nameof(MessageClass);
            public const string EntryId = nameof(MessageClass);
        }

        private IEnumerable<Outlook.Category> GetCategories(Outlook.Folder folder)
        {
            try
            {
                var filter = $"[{Columns.MessageClass}] = 'IPM.Configuration.CategoryList'";
                var table = folder.GetTable(filter, Outlook.OlTableContents.olHiddenItems);
                var columns = table.Columns;
                columns.RemoveAll();
                columns.Add(Columns.MessageClass);
                columns.Add(Columns.EntryId);
                columns.Add(Columns.Property);
                while (!table.EndOfTable)
                {
                    var row = table.GetNextRow();
                    var cls = row[Columns.MessageClass];
                    var id = row[Columns.EntryId];
                    var prop = row[Columns.Property];
                    var propS = Encoding.UTF8.GetString((byte[])prop);
                }

            }
            catch
            {

            }
            return new List<Outlook.Category>();
        }
    }
}
