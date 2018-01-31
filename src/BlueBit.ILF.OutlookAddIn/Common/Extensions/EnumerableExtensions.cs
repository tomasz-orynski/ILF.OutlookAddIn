using MoreLinq;
using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;

namespace BlueBit.ILF.OutlookAddIn.Common.Extensions
{
    public static class EnumerableExtensions
    {
        public static IEnumerable<T> NullAsEmpty<T>(this IEnumerable<T> @this) => @this != null ? @this : new T[] { };

        public static IEnumerable<(T Prev, T Curr, T Next)> AsNodes<T>(this IEnumerable<T> @this)
            where T : class
        {
            using (var enumerator = @this.GetEnumerator())
            {
                if (!enumerator.MoveNext()) yield break;
                T prev = null;
                while (true)
                {
                    var curr = enumerator.Current;
                    if (!enumerator.MoveNext())
                    {
                        yield return (prev, curr, null);
                        yield break;
                    }
                    yield return (prev, curr, enumerator.Current);
                    prev = curr;
                    curr = enumerator.Current;
                }
            }
        }

        public static IEnumerable<(T? Prev, T Curr, T? Next)> AsValueNodes<T>(this IEnumerable<T> @this)
            where T : struct
        {
            using (var enumerator = @this.GetEnumerator())
            {
                if (!enumerator.MoveNext()) yield break;
                T? prev = null;
                while (true)
                {
                    var curr = enumerator.Current;
                    if (!enumerator.MoveNext())
                    {
                        yield return (prev, curr, null);
                        yield break;
                    }
                    yield return (prev, curr, enumerator.Current);
                    prev = curr;
                    curr = enumerator.Current;
                }
            }
        }

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

        public static void ForEachFunc<T, TResult>(this IEnumerable<T> @this, Func<T, TResult> action)
            => @this.ForEach(_ => action(_));
    }
}