using System;
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

        public static Action<T> AsSingleEventHandler<T>(this Action<T> @this)
            => @param => AsSingleEventHandler(() => @this(@param));
        public static Action<T1,T2> AsSingleEventHandler<T1,T2>(this Action<T1, T2> @this)
            => (@param1, @param2) => AsSingleEventHandler(() => @this(@param1, @param2));

        public static Outlook.ItemsEvents_ItemAddEventHandler AsSingleItemAddEventHandler(this Action<object> @this)
            => new Outlook.ItemsEvents_ItemAddEventHandler(@this.AsSingleEventHandler());
    }
}
