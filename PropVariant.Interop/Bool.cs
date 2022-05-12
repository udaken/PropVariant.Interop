using System;
using System.Runtime.InteropServices;

namespace PropVarinatInterop
{
    [StructLayout(LayoutKind.Sequential)]
    public struct Bool : IEquatable<Bool>
    {
        public short Value;

        public static readonly Bool False = default;
        public static readonly Bool True = new Bool() { Value = -1 };

        public static implicit operator bool(Bool vb) => vb.Value != default;
        public bool Equals(Bool other)
        {
            return Value == other.Value;
        }

        public override bool Equals(object? obj) => obj is Bool b && Equals(b);

        public override int GetHashCode() => Value.GetHashCode();
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct Currency : IEquatable<Currency>
    {
        public long Value;

        public static implicit operator decimal(Currency cy) => decimal.FromOACurrency(cy.Value);
        public bool Equals(Currency other)
        {
            return Value == other.Value;
        }

        public override bool Equals(object? obj) => obj is Currency b && Equals(b);

        public override int GetHashCode() => Value.GetHashCode();
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct Date : IEquatable<Date>
    {
        public double Value;

        public static implicit operator DateTime(Date cy) => DateTime.FromOADate(cy.Value);
        public bool Equals(Date other)
        {
            return Value == other.Value;
        }

        public override bool Equals(object? obj) => obj is Date b && Equals(b);

        public override int GetHashCode() => Value.GetHashCode();
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct Error : IEquatable<Error>
    {
        public int Value;

        public bool Equals(Error other)
        {
            return Value == other.Value;
        }

        public override bool Equals(object? obj) => obj is Error b && Equals(b);

        public override int GetHashCode() => Value.GetHashCode();
    }
}


