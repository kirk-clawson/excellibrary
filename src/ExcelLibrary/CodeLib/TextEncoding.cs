using System;
using System.Collections.Generic;
using System.Text;

namespace QiHe.CodeLib
{
    /// <summary>
    /// TextEncoding
    /// </summary>
    public class TextEncoding
    {
        /// <summary>
        /// check if all charaters in text are ASCII charater
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public static bool FitsInASCIIEncoding(string text)
        {
            if(string.IsNullOrEmpty(text))return true;
            byte[] bytes = Encoding.UTF8.GetBytes(text);
            for (int i = 0; i < bytes.Length; i++)
            {
                if (bytes[i] > 127)
                {
                    return false;
                }
            }
            return true;
        }
    }
}
