using System.Diagnostics;

namespace ExcelLibrary.Infrastructure
{
    public static class BitMathExtensions
    {
        // 32 bit methdos
        public static uint GetBits(this uint source, int position, int width)
        {
            Debug.Assert(width > 0);
            Debug.Assert(position >= 0);
            Debug.Assert(position + width <= 32);

            var mask = (ulong) ((1 << width) - 1) << position;
            return (uint) (source & mask) >> position;
        }

        public static uint SetBits(this uint source, uint bitsToSet, int position, int width)
        {
            Debug.Assert(width > 0);
            Debug.Assert(position >= 0);
            Debug.Assert(position + width <= 32);

            var mask = (ulong) ((1 << width) - 1) << position;
            return (uint) (source & ~mask) | (bitsToSet << position);
        }

        public static bool GetBit(this uint source, int position)
        {
            Debug.Assert(position >= 0 && position < 32);

            return (source & (1 << position)) != 0;
        }

        public static uint SetBit(this uint source, bool value, int position)
        {
            Debug.Assert(position >= 0 && position < 32);

            var mask = (uint) (1 << position);
            if (value) return source | mask;
            return source & ~mask;
        }

        // 16 bit methods
        public static ushort GetBits(this ushort source, short position, short width)
        {
            Debug.Assert(width > 0);
            Debug.Assert(position >= 0);
            Debug.Assert(position + width <= 16);

            var mask = (uint) ((1 << width) - 1) << position;
            return (ushort) ((source & mask) >> position);
        }

        public static ushort SetBits(this ushort source, ushort bitsToSet, short position, short width)
        {
            Debug.Assert(width > 0);
            Debug.Assert(position >= 0);
            Debug.Assert(position + width <= 16);

            var mask = (uint) ((1 << width) - 1) << position;
            return (ushort) ((source & ~mask) | (ushort) (bitsToSet << position));
        }

        public static bool GetBit(this ushort source, short position)
        {
            Debug.Assert(position >= 0 && position < 16);

            return (source & (1 << position)) != 0;
        }

        public static ushort SetBit(this ushort source, bool value, short position)
        {
            Debug.Assert(position >= 0 && position < 16);

            var mask = (ushort) (1 << position);
            if (value) return (ushort) (source | mask);
            return (ushort) (source & ~mask);
        }

        // 8 bit methods
        public static byte GetBits(this byte source, byte position, byte width)
        {
            Debug.Assert(width > 0);
            Debug.Assert(position + width <= 8);

            var mask = (byte)((1 << width) - 1) << position;
            return (byte)((source & mask) >> position);
        }

        public static byte SetBits(this byte source, byte bitsToSet, byte position, byte width)
        {
            Debug.Assert(width > 0);
            Debug.Assert(position + width <= 8);

            var mask = (byte)((1 << width) - 1) << position;
            return (byte)((source & ~mask) | (byte)(bitsToSet << position));
        }

        public static bool GetBit(this byte source, byte position)
        {
            Debug.Assert(position < 8);

            return (source & (1 << position)) != 0;
        }

        public static byte SetBit(this byte source, bool value, byte position)
        {
            Debug.Assert(position < 8);

            var mask = (byte)(1 << position);
            if (value) return (byte)(source | mask);
            return (byte)(source & ~mask);
        }
    }
}