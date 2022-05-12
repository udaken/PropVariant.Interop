using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.Versioning;
using System.Diagnostics.CodeAnalysis;

namespace PropVarinatInterop
{
    [StructLayout(LayoutKind.Sequential)]
    public unsafe partial struct PROPVARIANT : IDisposable
    {
        public static bool PreferredI4ForInt32 { get; set; } = true;

        [StructLayout(LayoutKind.Explicit)]
        internal struct ValueUnion
        {
            [FieldOffset(0)]
            public byte UI1;
            [FieldOffset(0)]
            public sbyte I1;
            [FieldOffset(0)]
            public ushort UI2;
            [FieldOffset(0)]
            public short I2;
            [FieldOffset(0)]
            public uint UI4;
            [FieldOffset(0)]
            public int I4;
            [FieldOffset(0)]
            public ulong UI8;
            [FieldOffset(0)]
            public long I8;
            [FieldOffset(0)]
            public float R4;
            [FieldOffset(0)]
            public double R8;
            [FieldOffset(0)]
            public Bool Bool;
            [FieldOffset(0)]
            public Currency Currency;
            [FieldOffset(0)]
            public Date Date;
            [FieldOffset(0)]
            public Error Error;
            [FieldOffset(0)]
            public IntPtr IntPtr;
            [FieldOffset(0)]
            public long currency;
            [FieldOffset(0)]
            public CountedArray ca;
            [FieldOffset(0)]
            public Guid* Clsid;
            //[FieldOffset(0)]
            //public FILETIME filetime;

            [StructLayout(LayoutKind.Sequential)]
            public struct CountedArray
            {
                public nint cElems;
                public void* pElems;
            }
        };

        private VarType vt;
        private ushort reserved1;
        private ushort reserved2;
        private ushort reserved3;
        internal ValueUnion variantValue;

        public readonly ushort RawVt => (ushort)vt;

        public readonly VarType ElementType
        {
            get => (vt & VarType.TypeMask);
        }

        public readonly VarType VarType
        {
            get => vt;
        }
        private readonly VarTypeFlag Type
        {
            get => (VarTypeFlag)(vt & ~VarType.TypeMask);
        }
        enum VarTypeFlag
        {
            Scaler = 0x0000,
            Vector = VarType.Vector,
            Array = VarType.Array,
            ByRef = VarType.ByRef,
        }

        public readonly override string ToString()
        {
            return $"{(Type != VarTypeFlag.Scaler ? (Type + " | ") : Type)} {ElementType}";
        }

#if NETSTANDARD2_1_OR_GREATER || NETCOREAPP2_1_OR_GREATER
        public unsafe Span<T> AsSpan<T>() where T : unmanaged
        {
            switch (Type)
            {
                case VarTypeFlag.Scaler:
                    return MemoryMarshal.CreateSpan<T>(ref Unsafe.As<ValueUnion, T>(ref variantValue), 1);
                case VarTypeFlag.Vector:
                    return VectorAsSpan<T>();
                default:
                    throw new NotSupportedException();
            }
        }
#endif

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private readonly unsafe Span<T> VectorAsSpan<T>() where T : unmanaged
        {
            int count = (int)variantValue.ca.cElems;
            return new Span<T>(variantValue.ca.pElems, count);
        }

        private readonly unsafe T[] VectorAsArray<T>() where T : unmanaged
        {
            return VectorAsSpan<T>().ToArray();
        }

        public readonly bool TryReadDecimal(/*[NotNullWhen(true)] */ out decimal value)
        {
            switch (VarType)
            {
                case VarType.Decimal:
                    value = Unsafe.As<PROPVARIANT, decimal>(ref Unsafe.AsRef(in this));
                    return true;
                default:
                    value = default;
                    return false;
            }
        }
        public readonly bool TryReadString(/*[NotNullWhen(true)] */ out string? value)
        {
            switch (VarType)
            {
                case VarType.BStr:
                    value = Marshal.PtrToStringBSTR(variantValue.IntPtr);
                    break;
                case VarType.AnsiString:
                    value = Marshal.PtrToStringAnsi(variantValue.IntPtr);
                    break;
                case VarType.String:
                    value = Marshal.PtrToStringUni(variantValue.IntPtr);
                    break;
                default:
                    value = default;
                    return false;
            }
            return true;
        }
        public readonly bool TryReadGuid(out Guid value)
        {
            if (VarType is VarType.Clsid)
            {
                value = *variantValue.Clsid;
                return true;
            }
            else
            {
                value = default;
                return false;
            }
        }

        public readonly bool TryReadFileTime(out System.Runtime.InteropServices.ComTypes.FILETIME value)
        {
            if (VarType is VarType.FileTime)
            {
                value = Unsafe.As<long, System.Runtime.InteropServices.ComTypes.FILETIME>(ref Unsafe.AsRef(in variantValue.I8));
                return true;
            }
            else
            {
                value = default;
                return false;
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private readonly object? ReadScalerValue()
        {
            switch (VarType)
            {
                case VarType.Empty:
                case VarType.Null:
                    return null;
                case VarType.Blob:
                    return VectorAsArray<sbyte>();
                case VarType.I1:
                    return variantValue.I1;
                case VarType.UI1:
                    return variantValue.UI1;
                case VarType.I2:
                    return variantValue.I2;
                case VarType.UI2:
                    return variantValue.UI2;
                case VarType.UI4:
                case VarType.UInt:
                    return variantValue.UI4;
                case VarType.I4:
                case VarType.Int:
                    return variantValue.I4;
                case VarType.UI8:
                    return variantValue.UI8;
                case VarType.I8:
                    return variantValue.I8;
                case VarType.R4:
                    return variantValue.R4;
                case VarType.R8:
                    return variantValue.R8;
                case VarType.Currency:
                    return variantValue.Currency;
                case VarType.Bool:
                    return variantValue.Bool;
                case VarType.Date:
                    return variantValue.Date;
                case VarType.Error:
                    return variantValue.Error;
                case VarType.Unknown:
                case VarType.Dispatch:
                    return Marshal.GetObjectForIUnknown(variantValue.IntPtr);
                case VarType.IntPtr:
                case VarType.Ptr:
                    return variantValue.IntPtr;
                case VarType.UIntPtr:
                    return (UIntPtr)(void*)variantValue.IntPtr;
                case VarType.ClipboardFormat:
                    return Unsafe.AsRef<CLIPDATA>((void*)variantValue.IntPtr);
                case VarType.Clsid:
                    return Unsafe.AsRef<Guid>((void*)variantValue.IntPtr);
                case VarType.FileTime:
                    return DateTime.FromFileTime(variantValue.I8);
            }
            if (TryReadDecimal(out var decimalValue))
            {
                return decimalValue;
            }
            else if (TryReadString(out var stringValue))
            {
                return stringValue;
            }
            throw new NotSupportedException();
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private readonly unsafe object ReadVectorValue()
        {
            int count = (int)variantValue.ca.cElems;
            switch (VarType)
            {
                case VarType.VectorI1:
                    return VectorAsArray<sbyte>();
                case VarType.VectorUI1:
                    return VectorAsArray<byte>();
                case VarType.UI2:
                    return VectorAsArray<ushort>();
                case VarType.VectorI2:
                    return VectorAsArray<short>();
                case VarType.VectorUI4:
                    return VectorAsArray<uint>();
                case VarType.VectorI4:
                    return VectorAsArray<int>();
                case VarType.VectorUI8:
                    return VectorAsArray<ulong>();
                case VarType.VectorI8:
                    return VectorAsArray<long>();
                case VarType.VectorR4:
                    return VectorAsArray<float>();
                case VarType.VectorR8:
                    return VectorAsArray<double>();
                case VarType.VectorCurrency:
                    return VectorAsArray<Currency>();
                case VarType.VectorBool:
                    return VectorAsArray<Bool>();
                case VarType.VectorDate:
                    return VectorAsArray<Date>();
                case VarType.VectorError:
                    return VectorAsArray<Error>();
                case VarType.VectorBStr:
                    {
                        var array = new string?[count];
                        var span = VectorAsArray<IntPtr>();
                        for (int i = 0; i < count; i++)
                            array[i] = Marshal.PtrToStringBSTR(span[i]);
                        return array;
                    }
                case VarType.VectorAnsiString:
                    {
                        var array = new string?[count];
                        var span = VectorAsArray<IntPtr>();
                        for (int i = 0; i < count; i++)
                            array[i] = Marshal.PtrToStringAnsi(span[i]);
                        return array;
                    }
                case VarType.VectorString:
                    {
                        var array = new string?[count];
                        var span = VectorAsArray<IntPtr>();
                        for (int i = 0; i < count; i++)
                            array[i] = Marshal.PtrToStringUni(span[i]);
                        return array;
                    }
                case VarType.VectorClipboardFormat:
                    return VectorAsArray<CLIPDATA>();
                case VarType.VectorFileTime:
                    return VectorAsArray<System.Runtime.InteropServices.ComTypes.FILETIME>();
                case VarType.VectorClsid:
                    return VectorAsArray<Guid>();
                case VarType.VectorPropVariant:
                    return Marshal.GetObjectsForNativeVariants((IntPtr)variantValue.ca.pElems, count);
                default:
                    throw new NotSupportedException();
            }
        }

        public readonly unsafe object? Value
        {
            get
            {
                if (VarType == VarType.Illegal)
                    throw new ObjectDisposedException(nameof(PROPVARIANT));

                return Type switch
                {
                    VarTypeFlag.Scaler => ReadScalerValue(),
                    VarTypeFlag.Vector => ReadVectorValue(),
                    _ => throw new NotSupportedException(),
                };
            }
        }

        public void Clear()
        {
            PInvoke.PropVariantClear(ref this);
        }

        public void Dispose()
        {
            Clear();
            vt = VarType.Illegal;
        }

    }

}


