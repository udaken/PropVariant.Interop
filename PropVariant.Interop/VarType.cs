namespace PropVarinatInterop
{
    public enum VarType : ushort
    {
        /// <summary>
        /// nothing
        /// </summary>
        Empty = 0,
        /// <summary>
        /// SQL style Null
        /// </summary>
        Null = 1,
        /// <summary>
        /// 2 byte signed int
        /// </summary>
        I2 = 2,
        /// <summary>
        /// 4 byte signed int
        /// </summary>
        I4 = 3,
        /// <summary>
        /// 4 byte real
        /// </summary>
        R4 = 4,
        /// <summary>
        /// 8 byte real
        /// </summary>
        R8 = 5,
        /// <summary>
        /// currency
        /// </summary>
        Currency = 6,
        /// <summary>
        /// date
        /// </summary>
        Date = 7,
        /// <summary>
        /// OLE Automation string
        /// </summary>
        BStr = 8,
        /// <summary>
        /// IDispatch *
        /// </summary>
        Dispatch = 9,
        /// <summary>
        /// SCODE
        /// </summary>
        Error = 10,
        /// <summary>
        /// True=-1, False=0
        /// </summary>
        Bool = 11,
        /// <summary>
        /// VARIANT *
        /// </summary>
        Variant = 12,
        /// <summary>
        /// IUnknown *
        /// </summary>
        Unknown = 13,
        /// <summary>
        /// 16 byte fixed point
        /// </summary>
        Decimal = 14,

        I1 = 16,
        UI1 = 17,
        UI2 = 18,
        UI4 = 19,
        I8 = 20,
        UI8 = 21,
        Int = 22,
        UInt = 23,
        /// <summary>
        /// C style void
        /// </summary>
        Void = 24,
        /// <summary>
        /// Standard return type
        /// </summary>
        HResult = 25,
        /// <summary>
        /// pointer type
        /// </summary>
        Ptr = 26,
        /// <summary>
        /// (use VT_ARRAY in VARIANT)
        /// </summary>
        SafeArray = 27,
        /// <summary>
        /// C style array
        /// </summary>
        CStyleArray = 28,
        /// <summary>
        /// user defined type
        /// </summary>
        UserDefined = 29,
        /// <summary>
        /// null terminated string
        /// </summary>
        AnsiString = 30,
        /// <summary>
        /// wide null terminated string
        /// </summary>
        String = 31,
        /// <summary>
        /// user defined type
        /// </summary>
        Record = 36,
        /// <summary>
        /// signed machine register size width
        /// </summary>
        IntPtr = 37,
        /// <summary>
        /// unsigned machine register size width
        /// </summary>
        UIntPtr = 38,
        /// <summary>
        /// FILETIME
        /// </summary>
        FileTime = 64,
        /// <summary>
        /// Length prefixed bytes
        /// </summary>
        Blob = 65,
        /// <summary>
        /// Name of the stream follows
        /// </summary>
        Stream = 66,
        /// <summary>
        /// Name of the storage follows
        /// </summary>
        Storage = 67,
        /// <summary>
        /// Stream contains an object
        /// </summary>
        StreamedObject = 68,
        /// <summary>
        /// Storage contains an object
        /// </summary>
        StoredObject = 69,
        /// <summary>
        /// Blob contains an object 
        /// </summary>
        BlobObject = 70,
        /// <summary>
        /// Clipboard format
        /// </summary>
        ClipboardFormat = 71,
        /// <summary>
        /// A Class ID
        /// </summary>
        Clsid = 72,
        /// <summary>
        /// Stream with a GUID version
        /// </summary>
        VersionedStream = 73,
        /// <summary>
        /// Reserved for system use
        /// </summary>
        BstrBlob = 0xfff,

        ByRefI1 = ByRef | I1,
        ByRefI2 = ByRef | I2,
        ByRefI4 = ByRef | I4,
        ByRefI8 = ByRef | I8,
        ByRefUI1 = ByRef | UI1,
        ByRefUI2 = ByRef | UI2,
        ByRefUI4 = ByRef | UI4,
        ByRefUI8 = ByRef | UI8,
        ByRefInt = ByRef | Int,
        ByRefUInt = ByRef | UInt,
        ByRefR4 = ByRef | R4,
        ByRefR8 = ByRef | R8,
        ByRefBool = ByRef | Bool,
        ByRefDecimal = ByRef | Decimal,
        ByRefError = ByRef | Error,
        ByRefCurrency = ByRef | Currency,
        ByRefDate = ByRef | Date,
        ByRefBStr = ByRef | BStr,
        ByRefUnknown = ByRef | Unknown,
        ByRefDispatch = ByRef | Dispatch,
        ByRefSafeArray = ByRef | SafeArray,
        ByRefVariant = ByRef | Variant,

        VectorI1 = Vector | I1,
        VectorUI1 = Vector | UI1,
        VectorI2 = Vector | I2,
        VectorUI2 = Vector | UI2,
        VectorI4 = Vector | I4,
        VectorUI4 = Vector | UI4,
        VectorI8 = Vector | I8,
        VectorUI8 = Vector | UI8,
        VectorR4 = Vector | R4,
        VectorR8 = Vector | R8,
        VectorBool = Vector | Bool,
        VectorError = Vector | Error,
        VectorCurrency = Vector | Currency,
        VectorDate = Vector | Date,
        VectorFileTime = Vector | FileTime,
        VectorClsid = Vector | Clsid,
        VectorClipboardFormat = Vector | ClipboardFormat,
        VectorBStr = Vector | BStr,
        VectorBStrBlob = Vector | BstrBlob,
        VectorAnsiString = Vector | AnsiString,
        VectorString = Vector | String,
        VectorPropVariant = Vector | Variant,

        /// <summary>
        /// simple counted array
        /// </summary>
        Vector = 0x1000,
        /// <summary>
        /// SAFEARRAY*
        /// </summary>
        Array = 0x2000,
        /// <summary>
        /// void* for local use
        /// </summary>
        ByRef = 0x4000,
        Reserved = 0x8000,
        Illegal = 0xffff,
        TypeMask = 0x0fff,
    }
}
