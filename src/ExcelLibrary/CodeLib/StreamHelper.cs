using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace QiHe.CodeLib
{
    /// <summary>
    /// This class represents a Stream Helper.
    /// </summary>
    public class StreamHelper
    {
        /// <summary>
        /// Reads the bytes.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="count">The count.</param>
        /// <returns></returns>
        public static byte[] ReadBytes(Stream stream, int count)
        {
            int length = Math.Min((int)stream.Length, count);
            byte[] bytes = new byte[length];
            stream.Read(bytes, 0, bytes.Length);
            return bytes;
        }

        /// <summary>
        /// Reads to end.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns></returns>
        public static byte[] ReadToEnd(Stream stream)
        {
            byte[] bytes = new byte[stream.Length - stream.Position];
            stream.Read(bytes, 0, bytes.Length);
            return bytes;
        }
    }
}
