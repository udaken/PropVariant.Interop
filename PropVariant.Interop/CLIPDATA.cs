using System.Runtime.InteropServices;

namespace PropVarinatInterop
{
    [StructLayout(LayoutKind.Sequential)]
    public unsafe struct CLIPDATA
    {
        uint cbSize;
        int ulClipFmt;
        byte* pClipData;
    }

}


