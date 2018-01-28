using BlueBit.ILF.OutlookAddIn.Common.Utils;
using Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;

namespace BlueBit.ILF.OutlookAddIn.Components
{
    public interface IEnviroment
    {
        IReadOnlyList<(string id, string name)> GetCategories(MAPIFolder folder);
        IFoldersSource FoldersSource { get; }
    }
}
