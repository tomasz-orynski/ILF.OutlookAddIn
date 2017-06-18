using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;
using System.Windows;

namespace BlueBit.ILF.OutlookAddIn.Common.Extensions
{
    public static class FuncExtensions
    {
        public static Func<TModel,bool> ApplyParams<TModel,T>(this Func<TModel,T,bool> @this, T param)
        {
            Contract.Assert(@this != null);
            return model => @this(model, param);
        }
        public static Func<TModel, bool> ApplyParams<TModel, T1, T2>(this Func<TModel, T1, T2, bool> @this, T1 param1, T2 param2)
        {
            Contract.Assert(@this != null);
            return model => @this(model, param1, param2);
        }

        public static Action<TModel> IfTrueThenCloseWindow<TModel>(this Func<TModel, bool> @this, Window window)
        {
            Contract.Assert(@this != null);
            return model =>
            {
                if (@this(model))
                    window.Close();
            };
        }

        public static Func<T, TResult> Cast<T, TResult>(this Func<T, TResult> @this) => @this;

        public static Func<T, bool> AlwaysTrue<T>() => data => true;
        public static Func<T, bool> AlwaysFalse<T>() => data => false;
    }
}
