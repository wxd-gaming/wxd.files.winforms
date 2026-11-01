using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace wxd.files.winforms
{
    /// <summary>大文件切割工具：将超大文件按指定行数切割为多个小文件</summary>
    public static class FileSplitter
    {
        /// <summary>
        /// 切割文件
        /// </summary>
        /// <param name="filePath">原始文件路径</param>
        /// <param name="lineOffsets">行索引（每行的起始字节偏移数组）</param>
        /// <param name="fileLength">原始文件大小（字节）</param>
        /// <param name="linesPerFile">每个切割文件的行数</param>
        /// <param name="progressCallback">进度回调：(当前文件索引, 总文件数, 当前进度百分比)</param>
        /// <param name="headerLineCount">每个分片需要重复的顶部行数（CSV/TSV 表头传 1；0 表示不重复）</param>
        /// <returns>切割后的文件路径列表</returns>
        public static List<string> Split(string filePath, long[] lineOffsets, long fileLength,
            long linesPerFile, Action<int, int, int>? progressCallback = null, int headerLineCount = 0)
        {
            var result = new List<string>();
            if (linesPerFile <= 0 || lineOffsets.Length == 0) return result;
            string dir = Path.GetDirectoryName(filePath) ?? ".";
            string baseName = Path.GetFileNameWithoutExtension(filePath);
            string ext = Path.GetExtension(filePath);
            long totalLines = lineOffsets.Length;
            int fileCount = GetFileCount(totalLines, linesPerFile);
            int headerLines = Math.Max(0, Math.Min(headerLineCount, (int)Math.Min(totalLines, linesPerFile)));

            using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 65536, FileOptions.SequentialScan);

            for (int part = 0; part < fileCount; part++)
            {
                long startLine = part * linesPerFile;
                long endLine = Math.Min(startLine + linesPerFile, totalLines);
                long lineCount = endLine - startLine;

                string partPath = Path.Combine(dir, $"{baseName}-{part}{ext}");
                result.Add(partPath);

                // 计算本切割文件的字节范围
                long byteStart = lineOffsets[startLine];
                long byteEnd = (endLine < totalLines)
                    ? lineOffsets[endLine]
                    : fileLength;

                // 第 1 个分片自带表头，其余分片在顶部补一份，保证单独打开时表头可用
                using (var dest = new FileStream(partPath, FileMode.Create, FileAccess.Write, FileShare.None, 65536, FileOptions.SequentialScan))
                {
                    if (part > 0 && headerLines > 0)
                        CopyRange(fs, dest, lineOffsets[0], lineOffsets[headerLines]);
                    CopyRange(fs, dest, byteStart, byteEnd);
                }

                progressCallback?.Invoke(part + 1, fileCount, (int)((part + 1) * 100 / fileCount));
            }

            return result;
        }

        /// <summary>计算切割后的文件数量</summary>
        public static int GetFileCount(long totalLines, long linesPerFile)
        {
            if (linesPerFile <= 0) return 0;
            return (int)((totalLines + linesPerFile - 1) / linesPerFile);
        }

        /// <summary>流式拷贝文件的一个字节区间到目标流</summary>
        private static void CopyRange(FileStream source, FileStream dest, long byteStart, long byteEnd)
        {
            const int bufSize = 64 * 1024;
            byte[] buf = new byte[bufSize];
            long remaining = byteEnd - byteStart;

            source.Seek(byteStart, SeekOrigin.Begin);

            while (remaining > 0)
            {
                int toRead = (int)Math.Min(remaining, bufSize);
                int n = source.Read(buf, 0, toRead);
                if (n <= 0) break;
                dest.Write(buf, 0, n);
                remaining -= n;
            }
        }
    }
}
