using BlueBit.ILF.OutlookAddIn.Properties;
using NLog;
using System;
using System.Diagnostics.Contracts;
using System.Runtime.CompilerServices;
using System.Windows;

namespace BlueBit.ILF.OutlookAddIn.Diagnostics
{
    public static class LoggerExtensions
    {
        public static void EntryCall(Logger logger, Action action, string name)
        {
            try
            {
                logger.Trace(">>" + name);
                action();
                logger.Trace("<<" + name);
            }
            catch (Exception e)
            {
                logger.Error(e, "!!" + name);
                MessageBox.Show(e.Message, Resources.Exception_Caption, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        public static void EntryCall(Logger logger, Action action) => EntryCall(logger, action, action.Method.Name);

        public static void OnEntryCall(this Logger @this, Action action, [CallerMemberName]string name = null)
        {
            Contract.Assert(@this != null);
            Contract.Assert(action != null);
            Contract.Assert(!string.IsNullOrEmpty(name));
            EntryCall(@this, action, name);
        }

        public static T OnEntryCall<T>(this Logger @this, Func<T> action, [CallerMemberName]string name = null)
        {
            Contract.Assert(@this != null);
            Contract.Assert(action != null);
            Contract.Assert(!string.IsNullOrEmpty(name));
            var result = default(T);
            EntryCall(@this, () => result = action(), name);
            return result;
        }

        private static class _Lock<T>
        {
            public static volatile bool InCall = false;
        }

        public static void OnSingleEntryCall<TLock>(this Logger @this, Action action, [CallerMemberName]string name = null)
        {
            Contract.Assert(@this != null);
            Contract.Assert(action != null);
            Contract.Assert(!string.IsNullOrEmpty(name));

            if (_Lock<TLock>.InCall)
            {
                @this.Trace("~>" + name);
                return;
            }
            _Lock<TLock>.InCall = true;
            try
            {
                EntryCall(@this, action, name);
            }
            finally
            {
                _Lock<TLock>.InCall = false;
            }
        }

    }
}
