﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace BlueBit.ILF.OutlookAddIn.Common.Patterns
{
    public interface ICW<out T> : IDisposable
        where T : class
    {
        T Ref { get; }
    }

    internal sealed class CW<T> :
        DisposableBase,
        ICW<T>
        where T : class
    {
        private readonly T _ref;

        public T Ref => SafeCall(() => _ref);

        public CW(T @ref)
        {
            _ref = @ref;
        }

        protected override void OnDispose()
        {
            if (_ref != null)
            {
                var count = Marshal.FinalReleaseComObject(Ref);
            }
        }
    }

    public static class CWExt
    {
        public static ICW<T> AsCW<T>()
            where T : class, new()
            => new CW<T>(new T());

        public static ICW<T> AsCW<T>(this T @ref)
            where T : class
            => new CW<T>(@ref);
        public static ICW<T> AsCW_<T>(this object @ref)
            where T : class
            => new CW<T>((T)@ref);

        public static ICW<TProp> Call<T, TProp>(this ICW<T> @this, Func<T, TProp> getter)
            where T : class
            where TProp : class
            => getter(@this.Ref).AsCW();

        public static ICW<TProp> Call_<T, TProp>(this ICW<T> @this, Func<T, TProp> getter)
            where T : class
            where TProp : class
        {
            using (@this)
                return @this.Call(getter);
        }

        public static void ForEach<T, TItem>(this ICW<T> @this, Action<ICW<TItem>> action)
            where T : class, IEnumerable
            where TItem : class
        {
            foreach (TItem item in @this.Ref)
                using (var itemCw = item.AsCW())
                    action(itemCw);
        }
        public static void ForEach_<T, TItem>(this ICW<T> @this, Action<ICW<TItem>> action)
            where T : class, IEnumerable
            where TItem : class
        {
            using (@this)
                @this.ForEach(action);
        }
    }
}
