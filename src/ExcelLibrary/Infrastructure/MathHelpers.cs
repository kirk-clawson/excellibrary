using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLibrary.Infrastructure
{
    public static class MathHelpers
    {
        public static uint GetBits(this uint source, int position, int width)
        {
            Debug.Assert(width > 0);

            var mask = (uint)(1 << position);
            for (int i = 1; i < width - 1; i++)
            {
                mask += (uint)(1 << (position + i));
            }
            return (source & mask) >> position;
        }

        public static uint SetBits(this uint source, uint bitsToSet, int position, int width)
        {
            Debug.Assert(width > 0);

            var mask = (uint)(1 << position);
            for (int i = 1; i < width - 1; i++)
            {
                mask += (uint)(1 << (position + i));
            }
            return source & ~mask | bitsToSet << position;
        }

        public static ushort GetBits(this ushort source, byte position, byte width)
        {
            Debug.Assert(width > 0);

            var mask = (ushort)(1 << position);
            for (int i = 1; i < width - 1; i++)
            {
                mask += (ushort)(1 << (position + i));
            }
            return (ushort)((source & mask) >> position);
        }

        public static ushort SetBits(this ushort source, ushort bitsToSet, byte position, byte width)
        {
            Debug.Assert(width > 0);

            var mask = (ushort)(1 << position);
            for (int i = 1; i < width - 1; i++)
            {
                mask += (ushort)(1 << (position + i));
            }
            return (ushort)(source & ~mask | bitsToSet << position);
        }
    }
}
