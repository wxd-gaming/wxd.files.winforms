using System;
using System.Security.Cryptography;
using System.Text;

namespace wxd.files.winforms
{
    /// <summary>文本哈希：统一按 UTF-8 取字节，输出十六进制摘要。</summary>
    internal static class HashToolEngine
    {
        internal static string Md5Hex(string text, bool upper = false)
            => Format(MD5.HashData(Encoding.UTF8.GetBytes(text)), upper);

        internal static string Sha256Hex(string text, bool upper = false)
            => Format(SHA256.HashData(Encoding.UTF8.GetBytes(text)), upper);

        private static string Format(byte[] hash, bool upper)
        {
            string hex = Convert.ToHexString(hash);
            return upper ? hex : hex.ToLowerInvariant();
        }
    }
}
