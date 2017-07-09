using NLog;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.Contracts;
using System.Linq;

namespace BlueBit.ILF.OutlookAddIn.Diagnostics
{
    public static class EnumerableExtensions
    {
        private static Logger _logger = LogManager.GetCurrentClassLogger();

        public static List<T> SafeToList<T>(this IEnumerable<T> @this)
        {
            try { return @this.ToList(); }
            catch (Exception e)
            {
                var msg = $"{nameof(SafeToList)}<{typeof(T).Name}>";
                _logger.Warn(e, msg);
                if (Debugger.IsAttached)
                    Debugger.Break();
            }
            return new List<T>();
        }

        public static IEnumerable<T> SafeWhere<T>(this IEnumerable<T> @this, Func<T,bool> predicate)
        {
            Contract.Assert(predicate != null);
            return @this.Where(item =>
            {
                try { return predicate(item); }
                catch (Exception e)
                {
                    var msg = $"{nameof(SafeWhere)}<{typeof(T).Name}>";
                    _logger.Warn(e, msg);
                    if (Debugger.IsAttached)
                        Debugger.Break();
                }
                return false;
            });
        }
    }
}
