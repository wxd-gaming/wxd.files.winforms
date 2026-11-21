using System;
using System.Net;
using System.Text;

namespace wxd.files.winforms
{
    /// <summary>文本转换结果：成功时 Text 有值，失败时 Error 是给状态栏的中文提示。</summary>
    internal readonly struct TextToolResult
    {
        public bool Success { get; init; }
        public string Text { get; init; }
        public string? Error { get; init; }

        internal static TextToolResult Ok(string text) => new() { Success = true, Text = text };
        internal static TextToolResult Fail(string error) => new() { Success = false, Text = "", Error = error };
    }

    /// <summary>
    /// 文本编码工具：Unicode 转义/还原、UTF-8 Hex、Base64、URL、HTML。
    /// 文本与字节之间统一按 UTF-8；解码失败不返回替换字符，而是报错。
    /// </summary>
    internal static class EncodingToolEngine
    {
        private static readonly UTF8Encoding StrictUtf8 = new(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);

        // ==================== Unicode \uXXXX ====================

        /// <summary>文本转 \uXXXX。asciiOnly=true 时只转义非 ASCII 字符。</summary>
        internal static string EscapeUnicode(string text, bool asciiOnly)
        {
            var sb = new StringBuilder(text.Length * 2);
            foreach (char c in text)
            {
                if (asciiOnly && c <= 0x7F)
                    sb.Append(c);
                else
                    sb.Append("\\u").Append(((int)c).ToString("X4"));
            }
            return sb.ToString();
        }

        /// <summary>还原 \uXXXX / \UXXXXXXXX 和常见转义；不认识的转义按原样保留。</summary>
        internal static TextToolResult UnescapeUnicode(string text)
        {
            var sb = new StringBuilder(text.Length);
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (c != '\\' || i + 1 >= text.Length)
                {
                    sb.Append(c);
                    continue;
                }

                char next = text[i + 1];
                switch (next)
                {
                    case 'u':
                        if (!TryReadHex(text, i + 2, 4, out int code16))
                            return TextToolResult.Fail($"第 {i + 1} 个字符处的 \\u 转义不完整（需要 4 位十六进制）");
                        sb.Append((char)code16);
                        i += 5;
                        break;

                    case 'U':
                        if (!TryReadHex(text, i + 2, 8, out int code32))
                            return TextToolResult.Fail($"第 {i + 1} 个字符处的 \\U 转义不完整（需要 8 位十六进制）");
                        if (code32 > 0x10FFFF || (code32 >= 0xD800 && code32 <= 0xDFFF))
                            return TextToolResult.Fail($"第 {i + 1} 个字符处的 \\U 转义不是合法码点");
                        sb.Append(char.ConvertFromUtf32(code32));
                        i += 9;
                        break;

                    case 'n': sb.Append('\n'); i++; break;
                    case 'r': sb.Append('\r'); i++; break;
                    case 't': sb.Append('\t'); i++; break;
                    case '0': sb.Append('\0'); i++; break;
                    case '\\': sb.Append('\\'); i++; break;
                    case '"': sb.Append('"'); i++; break;
                    case '\'': sb.Append('\''); i++; break;
                    default:
                        sb.Append('\\');
                        break;
                }
            }
            return TextToolResult.Ok(sb.ToString());
        }

        private static bool TryReadHex(string text, int start, int count, out int value)
        {
            value = 0;
            if (start + count > text.Length) return false;
            for (int i = 0; i < count; i++)
            {
                int digit = HexDigit(text[start + i]);
                if (digit < 0) return false;
                value = value * 16 + digit;
            }
            return true;
        }

        // ==================== UTF-8 Hex ====================

        internal static string ToUtf8Hex(string text)
            => Convert.ToHexString(Encoding.UTF8.GetBytes(text)).Chunk(2)
                .Select(pair => new string(pair))
                .Aggregate(new StringBuilder(), (sb, s) => sb.Append(s).Append(' '))
                .ToString()
                .TrimEnd();

        internal static TextToolResult FromUtf8Hex(string hex)
        {
            if (!TryParseHexBytes(hex, out byte[] bytes, out string? error))
                return TextToolResult.Fail(error ?? "Hex 解析失败");

            try
            {
                return TextToolResult.Ok(StrictUtf8.GetString(bytes));
            }
            catch (DecoderFallbackException)
            {
                return TextToolResult.Fail("字节不是合法的 UTF-8 文本");
            }
        }

        private static bool TryParseHexBytes(string hex, out byte[] bytes, out string? error)
        {
            var cleaned = new StringBuilder(hex.Length);
            for (int i = 0; i < hex.Length; i++)
            {
                char c = hex[i];
                if (char.IsWhiteSpace(c) || c == ',' || c == ';' || c == '-' || c == ':' || c == '|')
                    continue;
                if (c == '0' && i + 1 < hex.Length && (hex[i + 1] == 'x' || hex[i + 1] == 'X'))
                {
                    i++;
                    continue;
                }
                if (HexDigit(c) < 0)
                {
                    bytes = Array.Empty<byte>();
                    error = $"包含非法十六进制字符“{c}”";
                    return false;
                }
                cleaned.Append(c);
            }

            if (cleaned.Length == 0)
            {
                bytes = Array.Empty<byte>();
                error = "没有可解析的十六进制内容";
                return false;
            }
            if (cleaned.Length % 2 != 0)
            {
                bytes = Array.Empty<byte>();
                error = "十六进制位数是奇数，无法按字节解析";
                return false;
            }

            bytes = new byte[cleaned.Length / 2];
            for (int i = 0; i < bytes.Length; i++)
                bytes[i] = (byte)(HexDigit(cleaned[i * 2]) * 16 + HexDigit(cleaned[i * 2 + 1]));
            error = null;
            return true;
        }

        private static int HexDigit(char c)
        {
            if (c >= '0' && c <= '9') return c - '0';
            if (c >= 'a' && c <= 'f') return c - 'a' + 10;
            if (c >= 'A' && c <= 'F') return c - 'A' + 10;
            return -1;
        }

        // ==================== Base64 ====================

        internal static string ToBase64(string text)
            => Convert.ToBase64String(Encoding.UTF8.GetBytes(text));

        internal static TextToolResult FromBase64(string base64)
        {
            var cleaned = new StringBuilder(base64.Length);
            foreach (char c in base64)
                if (!char.IsWhiteSpace(c)) cleaned.Append(c);

            string packed = cleaned.ToString();
            if (packed.Length == 0)
                return TextToolResult.Fail("没有可解码的 Base64 内容");
            if (packed.Length % 4 != 0)
                packed = packed.PadRight(packed.Length + (4 - packed.Length % 4), '=');

            try
            {
                byte[] bytes = Convert.FromBase64String(packed);
                return TextToolResult.Ok(StrictUtf8.GetString(bytes));
            }
            catch (FormatException)
            {
                return TextToolResult.Fail("不是合法的 Base64 文本");
            }
            catch (DecoderFallbackException)
            {
                return TextToolResult.Fail("解码出的字节不是合法 UTF-8 文本");
            }
        }

        // ==================== URL / HTML ====================

        internal static string UrlEncode(string text) => Uri.EscapeDataString(text);

        internal static string UrlDecode(string text)
        {
            try
            {
                return Uri.UnescapeDataString(text);
            }
            catch (UriFormatException ex)
            {
                return text;
            }
        }

        internal static string HtmlEncode(string text) => WebUtility.HtmlEncode(text);

        internal static string HtmlDecode(string text) => WebUtility.HtmlDecode(text);
    }
}
