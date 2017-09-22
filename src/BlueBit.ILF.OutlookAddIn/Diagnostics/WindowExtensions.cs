using NLog;
using System.Windows;

namespace BlueBit.ILF.OutlookAddIn.Diagnostics
{
    public static class WindowExtensions
    {
        public static void ShowDialog(this Window @this, Logger logger)
            => logger.OnEntryCall(() => @this.ShowDialog());
    }
}
