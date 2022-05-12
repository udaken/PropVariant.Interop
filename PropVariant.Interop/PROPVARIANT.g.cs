


using System;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.CompilerServices;

namespace PropVarinatInterop;
unsafe partial struct PROPVARIANT 
{
    public readonly Span< SByte > GetVectorSByte()
    {
        if(VarType is VarType.VectorI1
                        )
        {
            return VectorAsArray< SByte >();
        }
        else
        {
            throw new NotSupportedException();
        }
    }
    public readonly ref SByte GetByRefSByte()
    {
        if(VarType is VarType.ByRefI1
                        )
        {
            return ref Unsafe.AsRef< SByte >((void*)variantValue.IntPtr);
        }
        else
        {
            throw new NotSupportedException();
        }
    }
    public readonly bool TryReadSByte(out SByte value)
    {
        if(VarType is VarType.I1
                        )
        {
            value = variantValue.I1;
            return true;
        }
        else
        {
            value = default;
            return false;
        }
    }
}

unsafe partial struct PROPVARIANT 
{
    public readonly Span< Int16 > GetVectorInt16()
    {
        if(VarType is VarType.VectorI2
                        )
        {
            return VectorAsArray< Int16 >();
        }
        else
        {
            throw new NotSupportedException();
        }
    }
    public readonly ref Int16 GetByRefInt16()
    {
        if(VarType is VarType.ByRefI2
                        )
        {
            return ref Unsafe.AsRef< Int16 >((void*)variantValue.IntPtr);
        }
        else
        {
            throw new NotSupportedException();
        }
    }
    public readonly bool TryReadInt16(out Int16 value)
    {
        if(VarType is VarType.I2
                        )
        {
            value = variantValue.I2;
            return true;
        }
        else
        {
            value = default;
            return false;
        }
    }
}

unsafe partial struct PROPVARIANT 
{
    public readonly Span< Int32 > GetVectorInt32()
    {
        if(VarType is VarType.VectorI4
             or VarType.Int             )
        {
            return VectorAsArray< Int32 >();
        }
        else
        {
            throw new NotSupportedException();
        }
    }
    public readonly ref Int32 GetByRefInt32()
    {
        if(VarType is VarType.ByRefI4
             or VarType.Int             )
        {
            return ref Unsafe.AsRef< Int32 >((void*)variantValue.IntPtr);
        }
        else
        {
            throw new NotSupportedException();
        }
    }
    public readonly bool TryReadInt32(out Int32 value)
    {
        if(VarType is VarType.I4
             or VarType.Int             )
        {
            value = variantValue.I4;
            return true;
        }
        else
        {
            value = default;
            return false;
        }
    }
}

unsafe partial struct PROPVARIANT 
{
    public readonly Span< Int64 > GetVectorInt64()
    {
        if(VarType is VarType.VectorI8
                        )
        {
            return VectorAsArray< Int64 >();
        }
        else
        {
            throw new NotSupportedException();
        }
    }
    public readonly ref Int64 GetByRefInt64()
    {
        if(VarType is VarType.ByRefI8
                        )
        {
            return ref Unsafe.AsRef< Int64 >((void*)variantValue.IntPtr);
        }
        else
        {
            throw new NotSupportedException();
        }
    }
    public readonly bool TryReadInt64(out Int64 value)
    {
        if(VarType is VarType.I8
                        )
        {
            value = variantValue.I8;
            return true;
        }
        else
        {
            value = default;
            return false;
        }
    }
}

unsafe partial struct PROPVARIANT 
{
    public readonly Span< Byte > GetVectorByte()
    {
        if(VarType is VarType.VectorUI1
                        )
        {
            return VectorAsArray< Byte >();
        }
        else
        {
            throw new NotSupportedException();
        }
    }
    public readonly ref Byte GetByRefByte()
    {
        if(VarType is VarType.ByRefUI1
                        )
        {
            return ref Unsafe.AsRef< Byte >((void*)variantValue.IntPtr);
        }
        else
        {
            throw new NotSupportedException();
        }
    }
    public readonly bool TryReadByte(out Byte value)
    {
        if(VarType is VarType.UI1
                        )
        {
            value = variantValue.UI1;
            return true;
        }
        else
        {
            value = default;
            return false;
        }
    }
}

unsafe partial struct PROPVARIANT 
{
    public readonly Span< UInt16 > GetVectorUInt16()
    {
        if(VarType is VarType.VectorUI2
                        )
        {
            return VectorAsArray< UInt16 >();
        }
        else
        {
            throw new NotSupportedException();
        }
    }
    public readonly ref UInt16 GetByRefUInt16()
    {
        if(VarType is VarType.ByRefUI2
                        )
        {
            return ref Unsafe.AsRef< UInt16 >((void*)variantValue.IntPtr);
        }
        else
        {
            throw new NotSupportedException();
        }
    }
    public readonly bool TryReadUInt16(out UInt16 value)
    {
        if(VarType is VarType.UI2
                        )
        {
            value = variantValue.UI2;
            return true;
        }
        else
        {
            value = default;
            return false;
        }
    }
}

unsafe partial struct PROPVARIANT 
{
    public readonly Span< UInt32 > GetVectorUInt32()
    {
        if(VarType is VarType.VectorUI4
             or VarType.UInt             )
        {
            return VectorAsArray< UInt32 >();
        }
        else
        {
            throw new NotSupportedException();
        }
    }
    public readonly ref UInt32 GetByRefUInt32()
    {
        if(VarType is VarType.ByRefUI4
             or VarType.UInt             )
        {
            return ref Unsafe.AsRef< UInt32 >((void*)variantValue.IntPtr);
        }
        else
        {
            throw new NotSupportedException();
        }
    }
    public readonly bool TryReadUInt32(out UInt32 value)
    {
        if(VarType is VarType.UI4
             or VarType.UInt             )
        {
            value = variantValue.UI4;
            return true;
        }
        else
        {
            value = default;
            return false;
        }
    }
}

unsafe partial struct PROPVARIANT 
{
    public readonly Span< UInt64 > GetVectorUInt64()
    {
        if(VarType is VarType.VectorUI8
                        )
        {
            return VectorAsArray< UInt64 >();
        }
        else
        {
            throw new NotSupportedException();
        }
    }
    public readonly ref UInt64 GetByRefUInt64()
    {
        if(VarType is VarType.ByRefUI8
                        )
        {
            return ref Unsafe.AsRef< UInt64 >((void*)variantValue.IntPtr);
        }
        else
        {
            throw new NotSupportedException();
        }
    }
    public readonly bool TryReadUInt64(out UInt64 value)
    {
        if(VarType is VarType.UI8
                        )
        {
            value = variantValue.UI8;
            return true;
        }
        else
        {
            value = default;
            return false;
        }
    }
}

unsafe partial struct PROPVARIANT 
{
    public readonly Span< Single > GetVectorSingle()
    {
        if(VarType is VarType.VectorR4
                        )
        {
            return VectorAsArray< Single >();
        }
        else
        {
            throw new NotSupportedException();
        }
    }
    public readonly ref Single GetByRefSingle()
    {
        if(VarType is VarType.ByRefR4
                        )
        {
            return ref Unsafe.AsRef< Single >((void*)variantValue.IntPtr);
        }
        else
        {
            throw new NotSupportedException();
        }
    }
    public readonly bool TryReadSingle(out Single value)
    {
        if(VarType is VarType.R4
                        )
        {
            value = variantValue.R4;
            return true;
        }
        else
        {
            value = default;
            return false;
        }
    }
}

unsafe partial struct PROPVARIANT 
{
    public readonly Span< Double > GetVectorDouble()
    {
        if(VarType is VarType.VectorR8
                        )
        {
            return VectorAsArray< Double >();
        }
        else
        {
            throw new NotSupportedException();
        }
    }
    public readonly ref Double GetByRefDouble()
    {
        if(VarType is VarType.ByRefR8
                        )
        {
            return ref Unsafe.AsRef< Double >((void*)variantValue.IntPtr);
        }
        else
        {
            throw new NotSupportedException();
        }
    }
    public readonly bool TryReadDouble(out Double value)
    {
        if(VarType is VarType.R8
                        )
        {
            value = variantValue.R8;
            return true;
        }
        else
        {
            value = default;
            return false;
        }
    }
}

unsafe partial struct PROPVARIANT 
{
    public readonly Span< Bool > GetVectorBool()
    {
        if(VarType is VarType.VectorBool
                        )
        {
            return VectorAsArray< Bool >();
        }
        else
        {
            throw new NotSupportedException();
        }
    }
    public readonly ref Bool GetByRefBool()
    {
        if(VarType is VarType.ByRefBool
                        )
        {
            return ref Unsafe.AsRef< Bool >((void*)variantValue.IntPtr);
        }
        else
        {
            throw new NotSupportedException();
        }
    }
    public readonly bool TryReadBool(out Bool value)
    {
        if(VarType is VarType.Bool
                        )
        {
            value = variantValue.Bool;
            return true;
        }
        else
        {
            value = default;
            return false;
        }
    }
}

unsafe partial struct PROPVARIANT 
{
    public readonly Span< Currency > GetVectorCurrency()
    {
        if(VarType is VarType.VectorCurrency
                        )
        {
            return VectorAsArray< Currency >();
        }
        else
        {
            throw new NotSupportedException();
        }
    }
    public readonly ref Currency GetByRefCurrency()
    {
        if(VarType is VarType.ByRefCurrency
                        )
        {
            return ref Unsafe.AsRef< Currency >((void*)variantValue.IntPtr);
        }
        else
        {
            throw new NotSupportedException();
        }
    }
    public readonly bool TryReadCurrency(out Currency value)
    {
        if(VarType is VarType.Currency
                        )
        {
            value = variantValue.Currency;
            return true;
        }
        else
        {
            value = default;
            return false;
        }
    }
}

unsafe partial struct PROPVARIANT 
{
    public readonly Span< Date > GetVectorDate()
    {
        if(VarType is VarType.VectorDate
                        )
        {
            return VectorAsArray< Date >();
        }
        else
        {
            throw new NotSupportedException();
        }
    }
    public readonly ref Date GetByRefDate()
    {
        if(VarType is VarType.ByRefDate
                        )
        {
            return ref Unsafe.AsRef< Date >((void*)variantValue.IntPtr);
        }
        else
        {
            throw new NotSupportedException();
        }
    }
    public readonly bool TryReadDate(out Date value)
    {
        if(VarType is VarType.Date
                        )
        {
            value = variantValue.Date;
            return true;
        }
        else
        {
            value = default;
            return false;
        }
    }
}

unsafe partial struct PROPVARIANT 
{
    public readonly Span< Error > GetVectorError()
    {
        if(VarType is VarType.VectorError
                        )
        {
            return VectorAsArray< Error >();
        }
        else
        {
            throw new NotSupportedException();
        }
    }
    public readonly ref Error GetByRefError()
    {
        if(VarType is VarType.ByRefError
                        )
        {
            return ref Unsafe.AsRef< Error >((void*)variantValue.IntPtr);
        }
        else
        {
            throw new NotSupportedException();
        }
    }
    public readonly bool TryReadError(out Error value)
    {
        if(VarType is VarType.Error
                        )
        {
            value = variantValue.Error;
            return true;
        }
        else
        {
            value = default;
            return false;
        }
    }
}

unsafe partial struct PROPVARIANT 
{
    public readonly ref Decimal GetByRefDecimal()
    {
        if(VarType is VarType.ByRefDecimal
                        )
        {
            return ref Unsafe.AsRef< Decimal >((void*)variantValue.IntPtr);
        }
        else
        {
            throw new NotSupportedException();
        }
    }
}

unsafe partial struct PROPVARIANT 
{
    public readonly bool TryReadIntPtr(out IntPtr value)
    {
        if(VarType is VarType.IntPtr
                        )
        {
            value = variantValue.IntPtr;
            return true;
        }
        else
        {
            value = default;
            return false;
        }
    }
}

unsafe partial struct PROPVARIANT 
{
    public readonly Span< Guid > GetVectorGuid()
    {
        if(VarType is VarType.VectorClsid
                        )
        {
            return VectorAsArray< Guid >();
        }
        else
        {
            throw new NotSupportedException();
        }
    }
}

unsafe partial struct PROPVARIANT 
{
    public readonly Span< CLIPDATA > GetVectorCLIPDATA()
    {
        if(VarType is VarType.VectorClipboardFormat
                        )
        {
            return VectorAsArray< CLIPDATA >();
        }
        else
        {
            throw new NotSupportedException();
        }
    }
}

