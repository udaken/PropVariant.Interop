using System;
using System.Runtime.Versioning;

namespace PropVarinatInterop
{
    public sealed partial class PropertyVariant
    : System.Runtime.ConstrainedExecution.CriticalFinalizerObject, IDisposable
    {
        PROPVARIANT _Raw;
        bool _disposed = false;
        public ref PROPVARIANT Raw
        {
            get
            {
                if (_disposed)
                    throw new ObjectDisposedException(nameof(PropertyVariant));
                return ref _Raw;
            }
        }
        public void Dispose()
        {
            _Raw.Dispose();
            GC.SuppressFinalize(this);
            _disposed = true;
        }

        public PropertyVariant Clone()
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(PropertyVariant));

            var newobj = new PropertyVariant();
            PInvoke.PropVariantCopy(out newobj._Raw, _Raw);
            return newobj;
        }

    }

}


