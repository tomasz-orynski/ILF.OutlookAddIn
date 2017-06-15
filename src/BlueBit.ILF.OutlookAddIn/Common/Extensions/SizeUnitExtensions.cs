using System;

namespace BlueBit.ILF.OutlookAddIn.Common.Extensions
{
    public static class SizeUnitExtensions
    {
        static readonly string[] SizeSuffixes =
                           { "B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB" };
        static string Convert<T>(T value, int decimalPlaces, Func<T, (int mag, decimal adjustedSize)> convert)
            where T: struct, IEquatable<T>
        {
            if (value.Equals(default(T)))
                return "0.0 B";

            (var mag, var adjustedSize) = convert(value);
            if (Math.Round(adjustedSize, decimalPlaces) >= 1000)
            {
                mag += 1;
                adjustedSize /= 1024;
            }
            return string.Format("{0:n" + decimalPlaces + "} {1}",
                adjustedSize,
                SizeSuffixes[mag]);
        }

        static (int mag, decimal adjustedSize) Convert(decimal value, int mag)
            => (mag, value / (1L << (mag * 10)));

        static (int mag, decimal adjustedSize) Convert(int value)
            => Convert((decimal)value, (int)Math.Log(value, 1024));
        static (int mag, decimal adjustedSize) Convert(long value)
            => Convert((decimal)value, (int)Math.Log(value, 1024));

        public static string ToStringWithSizeUnit(this int @this, int decimalPlaces = 1)
            => Convert(@this, decimalPlaces, Convert);
        public static string ToStringWithSizeUnit(this long @this, int decimalPlaces = 1)
            => Convert(@this, decimalPlaces, Convert);
    }
}
