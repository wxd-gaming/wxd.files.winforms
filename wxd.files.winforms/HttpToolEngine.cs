using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace wxd.files.winforms
{
    /// <summary>
    /// HTTP 页的纯逻辑：键值对按 Content-Type 组装（URL 查询串 / JSON / 表单 / 原样拼接），
    /// 响应字节按 charset 解码（回退到面板选的编码），已保存请求的序列化往返。
    /// 发请求本身用自带 HttpClient，在面板里做。
    /// </summary>
    internal static class HttpToolEngine
    {
        static HttpToolEngine()
        {
            // 响应可能带 GBK 等旧编码的 charset，不注册代码页就解不出来
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        /// <summary>把键值对拼成查询串追加到 URL；URL 已带 ? 就用 & 续接。</summary>
        internal static string BuildUrl(string url, IReadOnlyList<KeyValuePair<string, string>> pairs)
        {
            if (pairs.Count == 0) return url;
            return url + (url.Contains('?') ? "&" : "?") + EncodePairs(pairs);
        }

        /// <summary>键值对组装成 JSON 对象，保持填写的顺序；同名键后者覆盖，值统一按字符串。</summary>
        internal static string BuildJsonBody(IReadOnlyList<KeyValuePair<string, string>> pairs)
        {
            var obj = new JsonObject();
            foreach (var kv in pairs)
                obj[kv.Key] = kv.Value;
            return obj.ToJsonString();
        }

        /// <summary>键值对组装成 x-www-form-urlencoded 请求体，按所选编码做 %XX 转义，同名键保留多份。</summary>
        internal static string BuildFormBody(IReadOnlyList<KeyValuePair<string, string>> pairs, Encoding encoding)
            => string.Join("&", pairs.Select(kv => UrlEncode(kv.Key, encoding) + "=" + UrlEncode(kv.Value, encoding)));

        /// <summary>自定义 Content-Type 的兜底：原样 k=v&amp;k=v 拼接不转义（游戏服常按原始串手写解析）。</summary>
        internal static string BuildRawBody(IReadOnlyList<KeyValuePair<string, string>> pairs)
            => string.Join("&", pairs.Select(kv => kv.Key + "=" + kv.Value));

        /// <summary>按名称取编码；未知或空名称回退 UTF-8（GBK/GB18030 需要代码页提供程序，静态构造里已注册）。</summary>
        internal static Encoding GetEncoding(string? name)
        {
            if (string.IsNullOrWhiteSpace(name)) return Encoding.UTF8;
            try { return Encoding.GetEncoding(name.Trim()); }
            catch (ArgumentException) { return Encoding.UTF8; }
        }

        /// <summary>响应实际采用的解码编码名：Content-Type 里的 charset 命中就用它，否则回退到面板选的编码。</summary>
        internal static string ResolveCharset(string? contentType, Encoding fallback)
            => ResolveEncoding(contentType, fallback).WebName;

        /// <summary>
        /// 响应字节按 Content-Type 的 charset 解码；charset 没给或认不出就按面板选的编码兜底，
        /// 图片/二进制内容返回 null（面板显示占位提示）。
        /// </summary>
        internal static string? DecodeBody(byte[] data, string? contentType, Encoding fallback)
        {
            if (data.Length == 0) return "";
            var media = contentType ?? "";
            if (media.StartsWith("image/") || media.StartsWith("audio/") || media.StartsWith("video/")
                || media.Contains("octet-stream"))
                return null;
            return ResolveEncoding(media, fallback).GetString(data);
        }

        private static string ExtractCharset(string media)
        {
            int index = media.IndexOf("charset=", StringComparison.OrdinalIgnoreCase);
            if (index < 0) return "";
            int start = index + "charset=".Length;
            int end = media.IndexOf(';', start);
            return (end < 0 ? media[start..] : media[start..end]).Trim().Trim('"');
        }

        private static Encoding ResolveEncoding(string? contentType, Encoding fallback)
        {
            string name = ExtractCharset(contentType ?? "");
            if (name.Length == 0) return fallback;
            try { return Encoding.GetEncoding(name); }
            catch (ArgumentException) { return fallback; }
        }

        private static string UrlEncode(string text, Encoding encoding)
        {
            var sb = new StringBuilder();
            foreach (byte b in encoding.GetBytes(text))
            {
                if ((b >= 'a' && b <= 'z') || (b >= 'A' && b <= 'Z') || (b >= '0' && b <= '9')
                    || b == '-' || b == '_' || b == '.' || b == '~')
                    sb.Append((char)b);
                else
                    sb.Append('%').Append(b.ToString("X2"));
            }
            return sb.ToString();
        }

        private static string EncodePairs(IReadOnlyList<KeyValuePair<string, string>> pairs)
            => string.Join("&", pairs.Select(kv =>
                Uri.EscapeDataString(kv.Key) + "=" + Uri.EscapeDataString(kv.Value)));

        internal static string SerializeRequests(IReadOnlyList<SavedHttpRequest> requests)
            => JsonSerializer.Serialize(requests, new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            });

        internal static List<SavedHttpRequest> DeserializeRequests(string json)
            => JsonSerializer.Deserialize<List<SavedHttpRequest>>(json) ?? new List<SavedHttpRequest>();
    }

    /// <summary>一条已保存的 HTTP 请求（名称 + 完整输入状态），存 .wxd-work/http-requests.json。</summary>
    internal sealed class SavedHttpRequest
    {
        public string Name { get; set; } = "";
        public string Method { get; set; } = "GET";
        public string Url { get; set; } = "";
        /// <summary>实际发送用的 Content-Type；BodyFormat 是旧版字段的兼容兜底（JSON/Form），新档不再依赖。</summary>
        public string? ContentType { get; set; }
        public string BodyFormat { get; set; } = "JSON";
        public string[][] Params { get; set; } = Array.Empty<string[]>();
        /// <summary>请求体键值（Body 页签）。旧档没有此字段：POST 老档的 Params 就是请求体。</summary>
        public string[][]? BodyParams { get; set; }
        public string[][] Headers { get; set; } = Array.Empty<string[]>();
        public int Count { get; set; } = 1;
        public int Threads { get; set; } = 1;
    }
}
