using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace wxd.files.winforms.RegressionTests;

/// <summary>
/// 日志分析行索引的纯逻辑回归：跨块边界行、分片反向查找、精确合并。
/// 背景：21GB 级日志“一直卡在索引中”，排查发现块边界 LF 恰好落在块尾时该行会被丢掉；
/// 旧反向查找每块无条件读满 64MB 回退窗口，127 个边界白读约 8GB。
/// </summary>
internal static partial class Program
{
    private static void RunLargeFileIndexTests()
    {
        Console.WriteLine("--- 日志分析行索引 ---");
        Run("索引 块边界落在LF后不丢行", Test_IndexBoundaryLineKept);
        Run("索引 块边界在行中不重复", Test_IndexBoundaryLineDeduped);
        Run("索引 反向找LF只读必要分片", Test_FindPrevNewlineSmallRead);
        Run("索引 反向找LF跨多分片", Test_FindPrevNewlineMultiChunk);
        Run("索引 反向找LF找不到返回下限", Test_FindPrevNewlineNotFound);
        Run("索引 合并去重与末尾虚拟行", Test_MergeDedupAndTrailing);
        Run("索引 多块拼接与逐行结果一致", Test_ScanMergeMatchesNaive);
    }

    /// <summary>临时文件工具：ASCII 写入后按流复用，测试结束删文件。</summary>
    private sealed class TempFile : IDisposable
    {
        public string Path { get; }
        public TempFile(string content)
        {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "wxd-index-test-" + Guid.NewGuid().ToString("N") + ".log");
            File.WriteAllBytes(Path, Encoding.ASCII.GetBytes(content));
        }
        public FileStream OpenRead() => new(Path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        public void Dispose() { try { File.Delete(Path); } catch { } }
    }

    private static long[] NaiveLineStarts(byte[] bytes)
    {
        var list = new List<long> { 0 };
        for (int i = 0; i < bytes.Length; i++)
            if (bytes[i] == (byte)'\n' && i + 1 < bytes.Length)
                list.Add(i + 1);
        return list.ToArray();
    }

    private static void Test_IndexBoundaryLineKept()
    {
        // "aaa\nbbb\nccc\n"：第一个 LF 在 3，块边界取 4（恰好是第二行行首）
        var content = "aaa\nbbb\nccc\n";
        using var tf = new TempFile(content);
        using var fs = tf.OpenRead();

        var s0 = LargeFileReaderControl.ScanBlock(fs, 0, 4, 3, addFirstLineStart: false, CancellationToken.None);
        Assert(s0 != null, "块0 应完成");
        long first1 = LargeFileReaderControl.FindPrevNewlineEndCore(fs, 4, 1024, 1024, null, out bool found1);
        Assert(found1 && first1 == 4, $"块1 反向查找应命中 LF@3 之后的位置 4，实际 {first1} found={found1}");
        var s1 = LargeFileReaderControl.ScanBlock(fs, first1, content.Length, 3, found1, CancellationToken.None);
        Assert(s1 != null, "块1 应完成");

        var merged = LargeFileReaderControl.MergeBlockLineStarts(
            new List<long>?[] { s0!.Value.LineStarts, s1!.Value.LineStarts }, content.Length);
        // 守的回归：上一块按 </<= blockEnd 判据不收该行，本块必须补上首行，否则 bbb 从结果里消失
        AssertSeq(new long[] { 0, 4, 8 }, merged, "块边界紧接 LF 时第二行不能被丢掉");
    }

    private static void Test_IndexBoundaryLineDeduped()
    {
        // 边界取 5（在 bbb 中间）：第二行行首 4 由块0 收录，块1 又补一次，合并必须去重
        var content = "aaa\nbbb\nccc\n";
        using var tf = new TempFile(content);
        using var fs = tf.OpenRead();

        var s0 = LargeFileReaderControl.ScanBlock(fs, 0, 5, 3, addFirstLineStart: false, CancellationToken.None);
        Assert(s0 != null, "块0 应完成");
        long first1 = LargeFileReaderControl.FindPrevNewlineEndCore(fs, 5, 1024, 1024, null, out bool found1);
        Assert(found1 && first1 == 4, $"块1 首行应回退到 4，实际 {first1}");
        var s1 = LargeFileReaderControl.ScanBlock(fs, first1, content.Length, 3, found1, CancellationToken.None);
        Assert(s1 != null, "块1 应完成");

        AssertSeq(new long[] { 4 }, s1!.Value.LineStarts.Take(1), "块1 首行补录 4");
        Assert(s0!.Value.LineStarts.Contains(4), "块0 也应收录行首 4");
        var merged = LargeFileReaderControl.MergeBlockLineStarts(
            new List<long>?[] { s0.Value.LineStarts, s1.Value.LineStarts }, content.Length);
        AssertSeq(new long[] { 0, 4, 8 }, merged, "重复行首只能保留一份");
    }

    private static void Test_FindPrevNewlineSmallRead()
    {
        // 每行 100 字节，position 落在第 300 行行首；最近 LF 就在 position-1
        var content = string.Concat(Enumerable.Repeat(new string('x', 99) + "\n", 500));
        using var tf = new TempFile(content);
        using var fs = tf.OpenRead();

        long position = 100L * 300;
        long read = 0;
        long got = LargeFileReaderControl.FindPrevNewlineEndCore(fs, position, 64 * 1024 * 1024, 64 * 1024,
            n => read += n, out bool found);
        Assert(found, "应找到上一个 LF");
        AssertEq(position, got, "LF 之后的偏移应等于 position");
        // 守的回归：老实现无条件读满 64MB 回退窗口，这里必须只读一个 64KB 分片以内
        Assert(read <= 64 * 1024, $"只应读一个分片（实际 {read} 字节）");
    }

    private static void Test_FindPrevNewlineMultiChunk()
    {
        // 一行 200KB 的超长行：position 取文件末尾，最近 LF 在 200KB 之外，必须继续往前翻
        var big = new string('y', 200 * 1024 - 1) + "\n";
        var content = "head\n" + big + new string('z', 200 * 1024);
        using var tf = new TempFile(content);
        using var fs = tf.OpenRead();

        long position = Encoding.ASCII.GetByteCount(content); // 末尾无 LF
        long read = 0;
        long got = LargeFileReaderControl.FindPrevNewlineEndCore(fs, position, 64 * 1024 * 1024, 64 * 1024,
            n => read += n, out bool found);
        Assert(found, "跨分片也应找到 LF");
        AssertEq(5L + Encoding.ASCII.GetByteCount(big), got, "返回超长行下一行的起始偏移");
        Assert(read > 64 * 1024, $"200KB 间距必须翻多个分片（实际 {read} 字节）");
    }

    private static void Test_FindPrevNewlineNotFound()
    {
        // 回退上限内没有 LF：返回回退下限，found=false（调用方据此不补首行）
        var content = new string('z', 4096) + "\n";
        using var tf = new TempFile(content);
        using var fs = tf.OpenRead();

        long got = LargeFileReaderControl.FindPrevNewlineEndCore(fs, 4096, 1024, 256, null, out bool found);
        Assert(!found, "上限内没有 LF 时 found 必须为 false");
        AssertEq(4096L - 1024, got, "找不到时从回退下限开始扫");
    }

    private static void Test_MergeDedupAndTrailing()
    {
        // 文件 10 字节、以 LF 结尾：最后一条行起始 == fileLength，属虚拟行要剪掉
        var perBlock = new List<long>?[] { new() { 4, 8 }, new() { 4, 9, 10 }, null };
        var merged = LargeFileReaderControl.MergeBlockLineStarts(perBlock, 10);
        AssertSeq(new long[] { 0, 4, 8, 9 }, merged, "重复偏移去重、虚拟行剪掉、缺块跳过");

        var noTrailing = LargeFileReaderControl.MergeBlockLineStarts(new List<long>?[] { new() { 4 }, new() { 7 } }, 9);
        AssertSeq(new long[] { 0, 4, 7 }, noTrailing, "末尾无 LF 时最后一行保留");
    }

    private static void Test_ScanMergeMatchesNaive()
    {
        // 随机行长 + 小分片 + 小块的端到端：合并结果必须与逐字节扫描完全一致
        var rnd = new Random(20261123);
        var sb = new StringBuilder();
        int lines = 400;
        for (int i = 0; i < lines; i++)
        {
            sb.Append(new string((char)('a' + i % 26), 1 + rnd.Next(40)));
            if (i < lines - 1 || rnd.Next(2) == 0) sb.Append('\n');
        }
        var content = sb.ToString();
        var raw = Encoding.ASCII.GetBytes(content);
        var expected = NaiveLineStarts(raw);

        foreach (int blockSize in new[] { 97, 512, 4096 })
        {
            using var tf = new TempFile(content);
            using var fs = tf.OpenRead();

            var perBlock = new List<long>?[(raw.Length + blockSize - 1) / blockSize];
            for (int b = 0; b < perBlock.Length; b++)
            {
                long start = (long)b * blockSize;
                long end = Math.Min(raw.Length, start + blockSize);
                long first = start;
                bool addFirst = false;
                if (b > 0)
                    first = LargeFileReaderControl.FindPrevNewlineEndCore(fs, start, 64 * 1024, 8192, null, out addFirst);
                var scan = LargeFileReaderControl.ScanBlock(fs, first, end, 37, addFirst, CancellationToken.None);
                Assert(scan != null, "扫描不应取消");
                perBlock[b] = scan!.Value.LineStarts;
            }

            var merged = LargeFileReaderControl.MergeBlockLineStarts(perBlock, raw.Length);
            AssertSeq(expected, merged, $"blockSize={blockSize} 合并结果应与逐字节扫描一致");
            for (int i = 1; i < merged.Length; i++)
                Assert(merged[i] > merged[i - 1], $"偏移必须严格递增：{merged[i - 1]} -> {merged[i]}");
        }
    }
}