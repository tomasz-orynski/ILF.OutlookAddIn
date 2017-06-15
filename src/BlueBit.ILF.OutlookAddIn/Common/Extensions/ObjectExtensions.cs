namespace BlueBit.ILF.OutlookAddIn.Common.Extensions
{
    public static class ObjectExtensions
    {
        public static T As<T>(this object @this)
            where T : class
            => @this as T;

        public interface IConvert<T>
        {
            TOther Cast<TOther>() where TOther : T;
        }

        private class _Convert<T> : IConvert<T>
        {
            public readonly T _source;
            public _Convert(T source) { _source = source; }
            public TOther Cast<TOther>() where TOther : T => (TOther)_source;
        }


        public static IConvert<T> Convert<T>(this T @this)
            => new _Convert<T>(@this);
    }
}
