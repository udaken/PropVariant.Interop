using System;
using System.Runtime.InteropServices;

namespace PropVarinatInterop
{
    [StructLayout(LayoutKind.Sequential)]
    public struct VersionedStream
    {
        public Guid guidVersion;
        public IntPtr pStream;
    }

}


