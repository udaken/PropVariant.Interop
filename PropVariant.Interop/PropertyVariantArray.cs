using System;

namespace PropVarinatInterop
{
    public unsafe sealed partial class PropertyVariantArray
        : System.Runtime.ConstrainedExecution.CriticalFinalizerObject, IDisposable
    {
        PROPVARIANT[]? _array = null;
        bool _disposed = false;

        public ReadOnlySpan<PROPVARIANT> AsSpan => _disposed ? throw new ObjectDisposedException(nameof(PropertyVariantArray)) : _array;

        public void Dispose()
        {
            PInvoke.FreePropVariantArray(_array);
            GC.SuppressFinalize(this);
            _disposed = true;
        }

    }

}


