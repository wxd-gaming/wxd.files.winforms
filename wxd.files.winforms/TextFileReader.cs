using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace wxd.files.winforms
{
    /// <summary>
    /// 文本文件读取：按行切分，自动识别 UTF-8 / GB18030。
    /// 中文项目的配置表经常是 GBK，直接按 UTF-8 读会整片乱码，比对结果就没有意义了。
    /// </summary>
    internal static class TextFileReader
    {
        private const int SampleBytes = 64 * 1024;
        private static bool _providerRegistered;

        /// <summary>读取整个文件文本（带编码识别）。</summary>
        internal static string ReadAllText(string path)
        {
            var encoding = DetectEncoding(path);
            using var reader = new StreamReader(path, encoding, detectEncodingFromByteOrderMarks: true);
            return reader.ReadToEnd();
        }

        /// <summary>读取文件全部行（不含行尾换行符）。</summary>
        internal static string[] ReadAllLines(string path)
        {
            var encoding = DetectEncoding(path);
            var lines = new List<string>();
            using var reader = new StreamReader(path, encoding, detectEncodingFromByteOrderMarks: true);
            string? line;
            while ((line = reader.ReadLine()) != null)
                lines.Add(line);
            return lines.ToArray();
        }

        /// <summary>
        /// 取文件头部 64KB 按严格 UTF-8 解码：成功就按 UTF-8，失败退回 GB18030（代码页 936）。
        /// 从尾部逐字节回退，避免正好切在多字节字符中间造成误判。
        /// </summary>
        internal static Encoding DetectEncoding(string path)
        {
            var utf8Strict = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);
            byte[] sample;
            int read;

            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                int size = (int)Math.Min(fs.Length, SampleBytes);
                sample = new byte[size];
                read = fs.Read(sample, 0, size);
            }

            for (int trim = 0; trim <= 3 && trim < read; trim++)
            {
                try
                {
                    utf8Strict.GetString(sample, 0, read - trim);
                    return new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
                }
                catch (DecoderFallbackException)
                {
                    // 尾部可能截断了多字节字符，少读一两个字节再试
                }
            }

            try
            {
                if (!_providerRegistered)
                {
                    Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                    _providerRegistered = true;
                }
                return Encoding.GetEncoding(936);
            }
            catch
            {
                return new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
            }
        }
    }
}
