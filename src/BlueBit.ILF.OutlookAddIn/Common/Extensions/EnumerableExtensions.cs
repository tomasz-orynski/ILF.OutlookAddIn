using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;

namespace BlueBit.ILF.OutlookAddIn.Common.Extensions
{
    public static class EnumerableExtensions
    {
        public static Func<string, bool> AsPrefixFilter(this IEnumerable<string> @this)
        {
            Contract.Assert(@this != null);
            if (@this.Any())
                return input =>
                {
                    input = input.ToLower();
                    return @this
                        .Select(_ => _.ToLower())
                        .Any(input.StartsWith);
                };
            return input => true;
        }
        public static Func<string, bool> AsEqualsFilter(this IEnumerable<string> @this)
        {
            Contract.Assert(@this != null);
            if (@this.Any())
                return input =>
                {
                    input = input.ToLower();
                    return @this
                        .Select(_ => _.ToLower())
                        .Any(input.Equals);
                };
            return input => true;
        }

        public static IEnumerable<T> DebugFetch<T>(this IEnumerable<T> @this)
#if DEBUG
        {
            try { return @this.ToList(); }
            catch {
                Contract.Assert(false);
                throw;
            }
        }
#else
            => @this;
#endif
    }
}
