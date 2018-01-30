using System;

namespace BlueBit.ILF.OutlookAddIn.Common.Patterns
{
    public abstract class DisposableBase : IDisposable
    {
        private bool _disposed = false;
        protected bool IsDisposed => _disposed;


        protected void Check()
        {
            if (_disposed)
                throw new ObjectDisposedException(GetType().Name);
        }

        protected T SafeCall<T>(Func<T> action)
        {
            Check();
            return action();
        }
        protected void SafeCall(Action action)
        {
            Check();
            action();
        }
        protected void SafeCall<T>(T param, Action<T> action)
        {
            Check();
            action(param);
        }

        public void Dispose()
        {
            if (_disposed) return;
            OnDispose();
            GC.SuppressFinalize(this);
            _disposed = true;
        }

        protected abstract void OnDispose();
    }

    public sealed class DisposableEvents : IDisposable
    {
        public event Action DisposeOccured;

        public void Dispose() => DisposeOccured?.Invoke();
    }
}