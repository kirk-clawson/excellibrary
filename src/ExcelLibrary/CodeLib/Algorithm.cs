using System;
using System.Collections;
using System.Collections.Generic;

namespace QiHe.CodeLib
{
    /// <summary>
    /// The Algorithm object is used to perform the calculating tasks.
    /// </summary>
    public class Algorithm
    {
        public static byte[] ArraySection(byte[] data, int index, int count)
        {
            byte[] section = new byte[count];
            Array.Copy(data, index, section, 0, count);
            return section;
        }
    }
}
