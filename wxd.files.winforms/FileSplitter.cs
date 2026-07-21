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
        /// <returns>切割后的文件路径列表</returns>
        public static List<string> Split(string filePath, long[] lineOffsets, long fileLength,
            long linesPerFile, Action<int, int, int>? progressCallback = null)
        {
            var result = new List<string>();
            string dir = Path.GetDirectoryName(filePath) ?? ".";
            string baseName = Path.GetFileNameWithoutExtension(filePath);
            string ext = Path.GetExtension(filePath);
            long totalLines = lineOffsets.Length;
            int fileCount = GetFileCount(totalLines, linesPerFile);

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

                // 流式拷贝，64KB 缓冲区
                WriteFileChunk(fs, partPath, byteStart, byteEnd);

                progressCallback?.Invoke(part + 1, fileCount, (int)((part + 1) * 100 / fileCount));
            }

            return result;
        }

        /// <summary>计算切割后的文件数量</summary>
        public static int GetFileCount(long totalLines, long linesPerFile)
        {
            return (int)((totalLines + linesPerFile - 1) / linesPerFile);
        }

        /// <summary>流式拷贝文件的一个字节区间到新文件</summary>
        private static void WriteFileChunk(FileStream source, string destPath, long byteStart, long byteEnd)
        {
            const int bufSize = 64 * 1024;
            byte[] buf = new byte[bufSize];
            long remaining = byteEnd - byteStart;

            source.Seek(byteStart, SeekOrigin.Begin);

            using var dest = new FileStream(destPath, FileMode.Create, FileAccess.Write, FileShare.None, bufSize, FileOptions.SequentialScan);

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
