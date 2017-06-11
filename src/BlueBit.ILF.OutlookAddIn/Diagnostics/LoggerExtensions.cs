using NLog;
using System;
using System.Diagnostics.Contracts;
using System.Runtime.CompilerServices;

namespace BlueBit.ILF.OutlookAddIn.Diagnostics
{
    public static class LoggerExtensions
    {
        private static void HandleEntryCall(Logger logger, Action action, string name = null)
        {
            try
            {
                logger.Trace(">>" + name);
                action();
                logger.Trace("<<" + name);
            }
            catch (Exception e)
            {
                logger.Error("!!" + name, e);
            }
        }

        public static void OnEntryCall(this Logger @this, Action action, [CallerMemberName]string name = null)
        {
            Contract.Assert(@this != null);
            Contract.Assert(action != null);
            Contract.Assert(!string.IsNullOrEmpty(name));
            HandleEntryCall(@this, action, name);
        }

        public static T OnEntryCall<T>(this Logger @this, Func<T> action, [CallerMemberName]string name = null)
        {
            Contract.Assert(@this != null);
            Contract.Assert(action != null);
            Contract.Assert(!string.IsNullOrEmpty(name));
            var result = default(T);
            HandleEntryCall(@this, () => result = action());
            return result;
        }
    }
}
