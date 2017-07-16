using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
