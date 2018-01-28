using Microsoft.Office.Interop.Outlook;
using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
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

        private static Logger _logger = LogManager.GetCurrentClassLogger();

        public static IEnumerable<(string id, string name)> GetCategories(this MAPIFolder folder)
        {
            var filter = $"[{Columns.MessageClass}] = 'IPM.Configuration.CategoryList'";
            var table = folder.GetTable(filter, OlTableContents.olHiddenItems);
            var columns = table.Columns;
            columns.RemoveAll();
            columns.Add(Columns.Property);

            var builder = new StringBuilder();
            while (!table.EndOfTable)
            {
                var row = table.GetNextRow();
                var prop = (byte[])row[Columns.Property];
                builder.Append(Encoding.UTF8.GetString(prop));
            }
            _logger.Trace(() => $"XML with categories [{folder.FolderPath}][{folder.Name}]:{Environment.NewLine}{builder.ToString()}");
            try
            {
                var xml = XDocument.Parse(builder.ToString());
                return xml
                    .Root //categories
                    .Elements() //category[]
                    .Select(_ => (
                        _.Attribute("guid").Value,
                        _.Attribute("name").Value
                        ))
                    ;
            }
            catch(System.Exception e)
            {
                _logger.Error(e);
                return null;
            }
        }
    }
}
