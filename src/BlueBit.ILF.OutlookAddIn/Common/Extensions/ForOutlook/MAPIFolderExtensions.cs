using Outlook = Microsoft.Office.Interop.Outlook;
using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Runtime.InteropServices;
using BlueBit.ILF.OutlookAddIn.Common.Patterns;

namespace BlueBit.ILF.OutlookAddIn.Common.Extensions.ForOutlook
{
    public static class MAPIFolderExtensions
    {
        static class Consts
        {
            public const string PropertyId = "http://schemas.microsoft.com/mapi/proptag/0x7C080102";
            public const string MessageClass = nameof(MessageClass);
            public const string MessageClassId = "IPM.Configuration.CategoryList";
        }

        private static Logger _logger = LogManager.GetCurrentClassLogger();

        public static IEnumerable<(string id, string name)> GetCategoriesFromStorage(this ICW<Outlook.MAPIFolder> folder)
        {
            var storage = folder.Ref
                .GetStorage(Consts.MessageClassId, Outlook.OlStorageIdentifierType.olIdentifyByMessageClass);
            var pa = storage.PropertyAccessor;
            var xmlStr = Encoding.ASCII.GetString((byte[])pa.GetProperty(Consts.PropertyId));
            _logger.Trace(() => $"XML with categories [{folder.Ref.FolderPath}][{folder.Ref.Name}]:{Environment.NewLine}{xmlStr}");
            return GetCategories(xmlStr);
        }

        public static IEnumerable<(string id, string name)> GetCategoriesFromTable(this ICW<Outlook.MAPIFolder> folder)
        {
            var builder = new StringBuilder();
            var filter = $"[{Consts.MessageClass}] = '{Consts.MessageClassId}'";
            using (var table = folder.Call(_ => _.GetTable(filter, Outlook.OlTableContents.olHiddenItems)))
            using (var columns = table.Call(_ => _.Columns))
            {
                columns.Ref.Add(Consts.PropertyId);
                while (!table.Ref.EndOfTable)
                {
                    using (var row = table.Call(_ => _.GetNextRow()))
                    {
                        var prop = (byte[])row.Ref[Consts.PropertyId];
                        builder.Append(Encoding.UTF8.GetString(prop));
                    }
                }
            }
            var xmlStr = builder.ToString();
            _logger.Trace(() => $"XML with categories [{folder.Ref.FolderPath}][{folder.Ref.Name}]:{Environment.NewLine}{xmlStr}");
            return GetCategories(xmlStr);
        }

        private static IEnumerable<(string id, string name)> GetCategories(string xmlStr)
        {
            if (string.IsNullOrWhiteSpace(xmlStr)) return null;
            var xml = XDocument.Parse(xmlStr);
            return xml
                .Root //categories
                .Elements() //category[]
                .Select(_ => (
                    _.Attribute("guid").Value,
                    _.Attribute("name").Value
                    ))
                ;
        }
    }
}
