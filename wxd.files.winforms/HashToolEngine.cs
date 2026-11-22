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

        /// <summary>流式计算文件 MD5：不整读进内存，GB 级文件也不会撑爆；失败抛 IO 异常。</summary>
        internal static string FileMd5Hex(string path, bool upper = false)
        {
            using var stream = File.OpenRead(path);
            return Format(MD5.HashData(stream), upper);
        }

        private static string Format(byte[] hash, bool upper)
        {
            string hex = Convert.ToHexString(hash);
            return upper ? hex : hex.ToLowerInvariant();
        }
    }
}
