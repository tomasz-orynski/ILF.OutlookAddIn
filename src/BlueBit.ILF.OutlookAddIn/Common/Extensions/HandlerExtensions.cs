using BlueBit.ILF.OutlookAddIn.Diagnostics;
using NLog;
using System;
using System.Linq.Expressions;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BlueBit.ILF.OutlookAddIn.Common.Extensions
{
    public static class HandlerExtensions
    {
        public static Action AsSingleEventHandler(this Action @this)
        {
            bool inCall = false;
            return () =>
            {
                if (inCall)
                    return;

                inCall = true;
                try
                {
                    @this();
                }
                finally
                {
                    inCall = false;
                }
            };
        }
    }
}
