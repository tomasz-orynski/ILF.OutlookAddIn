using System.Diagnostics;

namespace BlueBit.ILF.OutlookAddIn.Diagnostics
{
    public static class DebuggerExt
    {
        public static void BreakIfAttached()
        {
            if (Debugger.IsAttached)
                Debugger.Break();
        }
    }
}
