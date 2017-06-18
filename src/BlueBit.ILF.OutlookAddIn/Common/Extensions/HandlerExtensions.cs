using BlueBit.ILF.OutlookAddIn.Diagnostics;
using NLog;
using System;
using System.Linq.Expressions;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BlueBit.ILF.OutlookAddIn.Common.Extensions
{
    public static class HandlerExtensions
    {
        private class _HandlerRefT2<T1, T2>
        {
            public Func<T1, T2> Handler;
            public void Call(T1 p1, ref T2 p2) => p2 = Handler(p1);
        }

        public static Action<T> AsEntryCall<T>(this Action<T> action, Logger logger)
        {
            return @param => LoggerExtensions.EntryCall(
                logger, 
                () => action(@param),
                action.Method.Name);
        }
        public static Func<T,TResult> AsEntryCall<T,TResult>(this Func<T,TResult> action, Logger logger)
        {
            return @param =>
            {
                TResult result = default(TResult);
                LoggerExtensions.EntryCall(
                logger,
                () => result = action(@param),
                action.Method.Name);
                return result;
            };
        }

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

        public static Outlook.ApplicationEvents_11_ItemSendEventHandler AsItemSendHandler(this Func<object,bool> @this, Logger logger)
        {
            var handler = new _HandlerRefT2<object, bool>()
            {
                Handler = @this.AsEntryCall(logger)
            };
            return new Outlook.ApplicationEvents_11_ItemSendEventHandler(handler.Call);
        }


        public static Outlook.ItemsEvents_ItemAddEventHandler AsSingleItemAddEventHandler(this Action<object> @this, Logger logger)
            => new Outlook.ItemsEvents_ItemAddEventHandler(@this
                .AsEntryCall(logger)
                .AsSingleEventHandler());
    }
}
