using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace BlueBit.ILF.OutlookAddIn.Common.Extensions.ForOutlook
{
    public static class MAPIFolderExtensions
    {
        static class Columns
        {
            public const string Property = "http://schemas.microsoft.com/mapi/proptag/0x7C080102";
            public const string MessageClass = nameof(MessageClass);
            public const string EntryId = nameof(MessageClass);
        }

        public static IEnumerable<(string id, string name)> GetCategories(this MAPIFolder folder)
        {
            var filter = $"[{Columns.MessageClass}] = 'IPM.Configuration.CategoryList'";
            var table = folder.GetTable(filter, OlTableContents.olHiddenItems);
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
                return XDocument.Parse(Encoding.UTF8.GetString((byte[])row[Columns.Property]))
                    .Descendants("category")
                    .Select(_ => (
                        _.Attribute("guid").Value, 
                        _.Attribute("name").Value
                        ));
            }
            return new (string,string)[] {};
        }
    }
}
