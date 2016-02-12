using System.Diagnostics;

namespace ExcelLibrary.Infrastructure
{
    public static class BitMathExtensions
    {
        public static uint GetBits(this uint source, int position, int width)
        {
            Debug.Assert(width > 0);

            ulong mask = (ulong)((1 << width) - 1) << position;
            return (uint)(source & mask) >> position;
        }

        public static uint SetBits(this uint source, uint bitsToSet, int position, int width)
        {
            Debug.Assert(width > 0);

            ulong mask = (ulong)((1 << width) - 1) << position;
            return (uint)(source & ~mask) | (bitsToSet << position);
        }

        public static ushort GetBits(this ushort source, short position, short width)
        {
            Debug.Assert(width > 0);

            uint mask = (uint)((1 << width) - 1) << position;
            return (ushort)((source & mask) >> position);
        }

        public static ushort SetBits(this ushort source, ushort bitsToSet, short position, short width)
        {
            Debug.Assert(width > 0);

            uint mask = (uint)((1 << width) - 1) << position;
            return (ushort)((source & ~mask) | (bitsToSet << position));
        }

        public static bool GetBit(this uint source, int position)
        {
            return (source & (1 << position)) != 0;
        }

        public static uint SetBit(this uint source, bool value, int position)
        {
            return (source & (1 << position)) != 0;
        }
    }
}
