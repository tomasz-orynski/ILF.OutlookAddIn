using System;

namespace BlueBit.ILF.OutlookAddIn.Diagnostics
{
    public static class ObjectExtensions
    {
        public static bool SafeCheck<T>(this T @this, Func<T, bool> action, bool errorResult = false)
        {
            try
            {
                return action(@this);
            }
            catch
            {
                return errorResult;
            }
        }
    }
}
