using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using wxd.files.winforms;

namespace wxd.files.winforms.RegressionTests
{
    /// <summary>
    /// 零依赖回归 runner：验证最近修复中可纯逻辑/文件系统验证的部分。
    /// 退出码 0=全通过，1=有失败。UI 耦合项(#3/#4/#9a)与 svn 进程项(#9b)留人工验收。
    /// </summary>
    internal static class Program
    {
        private static int _passed;
        private static int _failed;
        private static int _skipped;

        private static int Main()
        {
            Console.WriteLine("=== wxd.files.winforms 回归测试 ===");

            // #1 ID 归并与排序共用比较器
            Run("#1 数字ID归并顺序+去重", Test_NumericMerge);
            Run("#1 02与2数值相等但字符串不同不误合并", Test_NumericTieBreak);
            Run("#1 非数字ID忽略大小写归并", Test_CaseInsensitiveMerge);
            Run("#1 空输入归并", Test_EmptyMerge);
            Run("#1 归并结果按共用比较器严格递增", Test_MergeStrictlyIncreasing);

            // #2 差异单元格 long 编码
            Run("#2 row>=65536 编码不碰撞", Test_EncodeNoOverflow);
            Run("#2 编码往返解码", Test_EncodeRoundTrip);
            Run("#2 HashSet<long> 跨65536边界去重", Test_EncodeHashSetBoundary);

            // #5/#6 Git log 解析
            Run("#5 subject 内部 | 保留", Test_GitSubjectPipe);
            Run("#6 坏日期跳过保留好行", Test_GitBadDateSkip);
            Run("#6 空输出/缺段安全", Test_GitEmptyAndShort);

            // #8 .git 文件/目录检测
            Run("#8 .git目录检测", Test_DetectGitDir);
            Run("#8 .git文件(worktree/子模块)检测", Test_DetectGitFile);
            Run("#8 .svn目录检测", Test_DetectSvn);

            // #9c 相对路径正斜杠
            Run("#9c 嵌套相对路径正斜杠", Test_RelativePathNested);
            Run("#9c 根级相对路径", Test_RelativePathRoot);

            // #3 ID 映射跳过计数
            Run("#3 空/重复ID跳过计数", Test_IdMapSkipsDuplicateAndEmpty);
            Run("#3 冻结行排除不计跳过", Test_IdMapExcludesFrozenRows);
            Run("#3 ID忽略大小写去重", Test_IdMapCaseInsensitive);
            Run("#3 列越界全部跳过", Test_IdMapIdColOutOfRange);

            // #4 + 发现A 对齐判定一致性
            Run("#4 对齐判定决策表(含越界回退)", Test_IsIdAlignUsableDecision);

            // #9a 同步滚动钳制
            Run("#9a 滚动索引钳制/空表保护", Test_ClampSyncScrollIndex);

            // #9b svn update 参数
            Run("#9b svn update 去尾分隔符防转义", Test_BuildSvnUpdateArguments);

            // 发现B 冲突行解析
            Run("冲突 SVN C标记识别", Test_ParseConflictLinesSvn);
            Run("冲突 Git CONFLICT/合并失败", Test_ParseConflictLinesGit);
            Run("冲突 正常输出不误报", Test_ParseConflictLinesNoFalsePositive);

            // #7 空文件导出（真实 git 集成，git 不可用则跳过）
            RunIntegration("#7 空文件导出成功", Test_EmptyFileExport);

            Console.WriteLine();
            Console.WriteLine($"通过 {_passed}  失败 {_failed}  跳过 {_skipped}");
            return _failed == 0 ? 0 : 1;
        }

        // ==================== #1 ====================

        private static void Test_NumericMerge()
        {
            var left = new List<string> { "1", "2", "3", "10" };
            var right = new List<string> { "2", "3", "4", "10", "11" };
            var cmp = VersionCompareForm.BuildIdComparison(left.Concat(right));
            var merged = VersionCompareForm.MergeOrderedIds(
                VersionCompareForm.SortIds(left, cmp), VersionCompareForm.SortIds(right, cmp), cmp);
            AssertSeq(new[] { "1", "2", "3", "4", "10", "11" }, merged, "数字ID应按数值归并去重");
        }

        private static void Test_NumericTieBreak()
        {
            var left = new List<string> { "2" };
            var right = new List<string> { "02" };
            var cmp = VersionCompareForm.BuildIdComparison(left.Concat(right));
            var merged = VersionCompareForm.MergeOrderedIds(
                VersionCompareForm.SortIds(left, cmp), VersionCompareForm.SortIds(right, cmp), cmp);
            // 数值相等但字符串不同：必须保留两条，不能误合并
            AssertSeq(new[] { "02", "2" }, merged, "02与2应视为不同ID");
        }

        private static void Test_CaseInsensitiveMerge()
        {
            var left = new List<string> { "b", "a", "C" };
            var right = new List<string> { "a", "d" };
            var cmp = VersionCompareForm.BuildIdComparison(left.Concat(right));
            var merged = VersionCompareForm.MergeOrderedIds(
                VersionCompareForm.SortIds(left, cmp), VersionCompareForm.SortIds(right, cmp), cmp);
            AssertSeq(new[] { "a", "b", "C", "d" }, merged, "非数字ID应忽略大小写归并去重");
        }

        private static void Test_EmptyMerge()
        {
            var empty = new List<string>();
            var one = new List<string> { "1" };
            var cmp = VersionCompareForm.BuildIdComparison(one);
            AssertSeq(Array.Empty<string>(),
                VersionCompareForm.MergeOrderedIds(empty, empty, cmp), "双空");
            AssertSeq(new[] { "1" },
                VersionCompareForm.MergeOrderedIds(one, empty, cmp), "左有右空");
            AssertSeq(new[] { "1" },
                VersionCompareForm.MergeOrderedIds(empty, one, cmp), "左空右有");
        }

        private static void Test_MergeStrictlyIncreasing()
        {
            var left = new List<string> { "5", "1", "3", "2", "4" };
            var right = new List<string> { "4", "6", "2", "8", "1" };
            var cmp = VersionCompareForm.BuildIdComparison(left.Concat(right));
            var merged = VersionCompareForm.MergeOrderedIds(
                VersionCompareForm.SortIds(left, cmp), VersionCompareForm.SortIds(right, cmp), cmp);
            AssertSeq(new[] { "1", "2", "3", "4", "5", "6", "8" }, merged, "归并全集");
            for (int i = 1; i < merged.Count; i++)
                Assert(cmp(merged[i - 1], merged[i]) < 0,
                    $"归并结果必须按共用比较器严格递增：{merged[i - 1]} 应 < {merged[i]}");
        }

        // ==================== #2 ====================

        private static void Test_EncodeNoOverflow()
        {
            Assert(VersionCompareForm.EncodeDiffCell(65536, 0) != VersionCompareForm.EncodeDiffCell(0, 0),
                "row=65536 与 row=0 编码必须不同（int 会回绕碰撞）");
            AssertEq(4294967296L, VersionCompareForm.EncodeDiffCell(65536, 0), "65536<<16 的 long 值");
            AssertEq(0L, VersionCompareForm.EncodeDiffCell(0, 0), "原点编码");
        }

        private static void Test_EncodeRoundTrip()
        {
            int[] rows = { 0, 1, 65535, 65536, 100000, 1 << 20 };
            int[] cols = { 0, 1, 255, 65535 };
            foreach (var r in rows)
                foreach (var c in cols)
                {
                    long key = VersionCompareForm.EncodeDiffCell(r, c);
                    int dr = (int)(key >> 16);
                    int dc = (int)(key & 0xFFFF);
                    AssertEq(r, dr, $"行往返 row={r} col={c}");
                    AssertEq(c, dc, $"列往返 row={r} col={c}");
                }
        }

        private static void Test_EncodeHashSetBoundary()
        {
            var set = new HashSet<long>
            {
                VersionCompareForm.EncodeDiffCell(0, 0),
                VersionCompareForm.EncodeDiffCell(65536, 0),
                VersionCompareForm.EncodeDiffCell(65535, 65535),
                VersionCompareForm.EncodeDiffCell(65536, 0), // 重复，应去重
            };
            AssertEq(3, set.Count, "跨 65536 边界的三个不同单元格 + 一个重复");
            Assert(set.Contains(VersionCompareForm.EncodeDiffCell(65536, 0)), "应含 65536 行编码");
        }

        // ==================== #5 / #6 ====================

        private static void Test_GitSubjectPipe()
        {
            var output = "fullsha1|sha1|2026-01-02 03:04:05 +0800|Alice|fix: a|b|c subject\n";
            var entries = VersionControlHelper.ParseGitLogOutput(output);
            AssertEq(1, entries.Count, "应解析出 1 条");
            AssertEq("fix: a|b|c subject", entries[0].Message, "subject 内部 | 必须保留");
            AssertEq("sha1", entries[0].Revision, "短SHA");
            AssertEq("fullsha1", entries[0].FullRevision, "完整SHA");
            AssertEq("Alice", entries[0].Author, "作者");
        }

        private static void Test_GitBadDateSkip()
        {
            var output =
                "full1|sha1|2026-01-01 00:00:00 +0000|Bob|good one\n" +
                "full2|sha2|NOT-A-DATE|Bob|bad date\n" +
                "full3|sha3|2026-03-03 00:00:00 +0000|Bob|good two\n";
            var entries = VersionControlHelper.ParseGitLogOutput(output);
            AssertEq(2, entries.Count, "坏日期行跳过，好行保留（原实现会毁掉整段历史）");
            AssertEq("good one", entries[0].Message, "第一条好行");
            AssertEq("good two", entries[1].Message, "第二条好行");
            AssertEq(0, entries[0].Index, "重排序号 0");
            AssertEq(1, entries[1].Index, "重排序号 1");
        }

        private static void Test_GitEmptyAndShort()
        {
            AssertEq(0, VersionControlHelper.ParseGitLogOutput("").Count, "空输出");
            AssertEq(0, VersionControlHelper.ParseGitLogOutput("\n\n").Count, "仅空行");
            AssertEq(0, VersionControlHelper.ParseGitLogOutput("incomplete|line\n").Count, "缺段行跳过");
            AssertEq(1, VersionControlHelper.ParseGitLogOutput("f|s|2026-01-01 00:00:00 +0000|A|m\n").Count,
                "正常行仍解析");
        }

        // ==================== #8 ====================

        private static void Test_DetectGitDir()
        {
            var root = NewTempDir();
            try
            {
                Directory.CreateDirectory(Path.Combine(root, ".git"));
                var file = MakeFile(root, "sub", "data.csv");
                AssertEq(VcsType.Git, VersionControlHelper.DetectVcs(file), "DetectVcs(.git目录)");
                AssertEq(root, VersionControlHelper.GetVcsRoot(file, VcsType.Git), "GetVcsRoot(.git目录)");
            }
            finally { Directory.Delete(root, true); }
        }

        private static void Test_DetectGitFile()
        {
            var root = NewTempDir();
            try
            {
                // worktree/子模块：.git 是文件而非目录
                File.WriteAllText(Path.Combine(root, ".git"), "gitdir: /some/where\n");
                var file = MakeFile(root, "sub", "data.csv");
                AssertEq(VcsType.Git, VersionControlHelper.DetectVcs(file), "DetectVcs(.git文件)");
                AssertEq(root, VersionControlHelper.GetVcsRoot(file, VcsType.Git), "GetVcsRoot(.git文件)");
            }
            finally { Directory.Delete(root, true); }
        }

        private static void Test_DetectSvn()
        {
            var root = NewTempDir();
            try
            {
                Directory.CreateDirectory(Path.Combine(root, ".svn"));
                var file = MakeFile(root, "data.csv");
                AssertEq(VcsType.Svn, VersionControlHelper.DetectVcs(file), "DetectVcs(.svn目录)");
                AssertEq(root, VersionControlHelper.GetVcsRoot(file, VcsType.Svn), "GetVcsRoot(.svn目录)");
            }
            finally { Directory.Delete(root, true); }
        }

        // ==================== #9c ====================

        private static void Test_RelativePathNested()
        {
            var root = NewTempDir();
            try
            {
                Directory.CreateDirectory(Path.Combine(root, ".git"));
                var file = MakeFile(root, "sub", "dir", "data.csv");
                var rel = VersionControlHelper.GetRelativePath(file, VcsType.Git);
                AssertEq("sub/dir/data.csv", rel, "嵌套相对路径");
                Assert(!rel.Contains('\\'), "相对路径不应含反斜杠");
            }
            finally { Directory.Delete(root, true); }
        }

        private static void Test_RelativePathRoot()
        {
            var root = NewTempDir();
            try
            {
                Directory.CreateDirectory(Path.Combine(root, ".git"));
                var file = MakeFile(root, "data.csv");
                AssertEq("data.csv", VersionControlHelper.GetRelativePath(file, VcsType.Git), "根级相对路径");
            }
            finally { Directory.Delete(root, true); }
        }

        // ==================== #3 ID 映射跳过计数 ====================

        private static void Test_IdMapSkipsDuplicateAndEmpty()
        {
            var fd = MakeFd(
                (0, new[] { "1", "a" }),
                (1, new[] { "2", "b" }),
                (2, new[] { "1", "c" }),   // 重复 ID
                (3, new[] { "", "d" }),    // 空 ID
                (4, new[] { "  ", "e" }),  // 纯空白 ID
                (5, new[] { "3", "f" }));
            var (map, skipped) = VersionCompareForm.BuildIdRowMap(fd, 0, 0);
            AssertEq(3, map.Count, "有效 ID 数");
            AssertEq(3, skipped, "重复1+空1+空白1=3");
            AssertEq(0, map["1"], "ID 1 → 行索引 0");
            AssertEq(1, map["2"], "ID 2 → 行索引 1");
            AssertEq(5, map["3"], "ID 3 → 行索引 5");
        }

        private static void Test_IdMapExcludesFrozenRows()
        {
            var fd = MakeFd(
                (0, new[] { "H1" }),
                (1, new[] { "H2" }),
                (2, new[] { "1" }),
                (3, new[] { "2" }),
                (4, new[] { "1" }));  // 重复
            var (map, skipped) = VersionCompareForm.BuildIdRowMap(fd, 2, 0);
            AssertEq(2, map.Count, "冻结行排除后有效 ID");
            AssertEq(1, skipped, "仅重复行计跳过，冻结行不计");
            AssertEq(2, map["1"], "ID 1 → fd 行索引 2");
            AssertEq(3, map["2"], "ID 2 → fd 行索引 3");
        }

        private static void Test_IdMapCaseInsensitive()
        {
            var fd = MakeFd(
                (0, new[] { "Abc" }),
                (1, new[] { "abc" }),  // 大小写重复
                (2, new[] { "XYZ" }));
            var (map, skipped) = VersionCompareForm.BuildIdRowMap(fd, 0, 0);
            AssertEq(2, map.Count, "忽略大小写去重");
            AssertEq(1, skipped, "abc 视为 Abc 重复");
            Assert(map.ContainsKey("ABC"), "映射应忽略大小写");
        }

        private static void Test_IdMapIdColOutOfRange()
        {
            var fd = MakeFd(
                (0, new[] { "1", "a" }),
                (1, new[] { "2", "b" }));
            var (map, skipped) = VersionCompareForm.BuildIdRowMap(fd, 0, 5);
            AssertEq(0, map.Count, "列越界→无有效 ID");
            AssertEq(2, skipped, "列越界→全部跳过");
        }

        // ==================== #4 + 发现A 对齐判定 ====================

        private static void Test_IsIdAlignUsableDecision()
        {
            Assert(!VersionCompareForm.IsIdAlignUsable(false, 0, 3, 3), "#4 取消勾选→回退行号对齐");
            Assert(VersionCompareForm.IsIdAlignUsable(true, 0, 3, 3), "有效列 0");
            Assert(VersionCompareForm.IsIdAlignUsable(true, 2, 3, 3), "有效末列 2");
            Assert(!VersionCompareForm.IsIdAlignUsable(true, -1, 3, 3), "负索引");
            Assert(!VersionCompareForm.IsIdAlignUsable(true, 3, 3, 3), "等于左列数→越界");
            Assert(!VersionCompareForm.IsIdAlignUsable(true, 2, 3, 2), "等于右列数→越界(发现A)");
            Assert(!VersionCompareForm.IsIdAlignUsable(true, 5, 3, 3), "远超列数");
        }

        // ==================== #9a 同步滚动钳制 ====================

        private static void Test_ClampSyncScrollIndex()
        {
            AssertEq((int?)null, VersionCompareForm.ClampSyncScrollIndex(0, 0), "空目标表不同步");
            AssertEq((int?)null, VersionCompareForm.ClampSyncScrollIndex(5, 0), "空目标表不同步 2");
            AssertEq((int?)null, VersionCompareForm.ClampSyncScrollIndex(-1, 10), "无效源索引");
            AssertEq((int?)3, VersionCompareForm.ClampSyncScrollIndex(3, 10), "正常索引");
            AssertEq((int?)9, VersionCompareForm.ClampSyncScrollIndex(10, 10), "idx==rowCount 钳到 9");
            AssertEq((int?)9, VersionCompareForm.ClampSyncScrollIndex(100, 10), "越界钳到 9");
            AssertEq((int?)0, VersionCompareForm.ClampSyncScrollIndex(0, 1), "单行表");
            AssertEq((int?)null, VersionCompareForm.ClampSyncScrollIndex(5, -1), "负 rowCount 防御");
        }

        // ==================== #9b svn update 参数 ====================

        private static void Test_BuildSvnUpdateArguments()
        {
            var normal = VersionControlHelper.BuildSvnUpdateArguments(@"C:\foo");
            AssertEq("update \"C:\\foo\" --non-interactive", normal, "正常路径参数");

            var trailing = VersionControlHelper.BuildSvnUpdateArguments(@"C:\foo\");
            AssertEq("update \"C:\\foo\" --non-interactive", trailing, "尾分隔符应去除");

            var spaces = VersionControlHelper.BuildSvnUpdateArguments(@"C:\my dir\");
            AssertEq("update \"C:\\my dir\" --non-interactive", spaces, "含空格路径去尾分隔符");

            // 关键回归：闭合引号前不能是反斜杠，否则 Windows argv 把 \" 当转义引号
            Assert(!trailing.Contains("\\\""), "参数不应含 \\\" 转义序列");
            Assert(!normal.Contains("\\\""), "正常参数不应含 \\\" 转义序列");
        }

        // ==================== 发现B 冲突行解析 ====================

        private static void Test_ParseConflictLinesSvn()
        {
            var combined = "U    file1.txt\nC    file2.txt\nA    file3.txt\nUpdated to revision 5.\n";
            var lines = VersionControlHelper.ParseConflictLines(combined);
            AssertEq(1, lines.Count, "只识别 C 标记行");
            AssertEq("C    file2.txt", lines[0], "SVN 冲突行");
        }

        private static void Test_ParseConflictLinesGit()
        {
            var combined = "Auto-merging file.txt\nCONFLICT (content): Merge conflict in file.txt\n" +
                           "Automatic merge failed; fix conflicts and then commit the result.\n";
            var lines = VersionControlHelper.ParseConflictLines(combined);
            AssertEq(2, lines.Count, "CONFLICT + Automatic merge failed");
            Assert(lines.Any(l => l.StartsWith("CONFLICT")), "含 CONFLICT 行");
            Assert(lines.Any(l => l.Contains("Automatic merge failed")), "含合并失败行");
        }

        private static void Test_ParseConflictLinesNoFalsePositive()
        {
            var combined = "Already up to date.\nno conflicts\nChecked out revision 10.\nSummary: 0 conflicts\n";
            AssertEq(0, VersionControlHelper.ParseConflictLines(combined).Count, "正常输出不应误报冲突");
            AssertEq(1, VersionControlHelper.ParseConflictLines("Summary of conflicts:\n   Text conflicts: 1\n").Count,
                "Summary of conflicts 应识别");
            AssertEq(1, VersionControlHelper.ParseConflictLines("Conflict discovered in 'file.txt'.\n").Count,
                "Conflict discovered 应识别");
        }

        // ==================== #7（真实 git 集成） ====================

        private static void Test_EmptyFileExport()
        {
            if (!GitAvailable())
                throw new SkipException("git 不可用");

            var repo = NewTempDir();
            var outDir = NewTempDir();
            try
            {
                Git(repo, "init");
                var file = Path.Combine(repo, "empty.txt");
                File.WriteAllText(file, "");
                Git(repo, "add empty.txt");
                Git(repo, "-c user.email=reg@test.local -c user.name=reg commit -m init");

                var result = VersionControlHelper.ExportRevisionAsync(file, VcsType.Git, "HEAD", outDir)
                    .GetAwaiter().GetResult();
                Assert(result.Success, "空文件导出应成功（原实现误判失败）: " + result.ErrorMessage);
                Assert(result.TempFile != null && File.Exists(result.TempFile), "应生成临时文件");
                AssertEq(0L, new FileInfo(result.TempFile!).Length, "导出文件应为空");
            }
            finally
            {
                DeleteTempDir(repo);
                DeleteTempDir(outDir);
            }
        }

        // ==================== 基础设施 ====================

        private static void Run(string name, Action test)
        {
            try
            {
                test();
                _passed++;
                Console.WriteLine($"  PASS  {name}");
            }
            catch (Exception ex)
            {
                _failed++;
                Console.WriteLine($"  FAIL  {name}: {ex.Message}");
            }
        }

        private static void RunIntegration(string name, Action test)
        {
            try
            {
                test();
                _passed++;
                Console.WriteLine($"  PASS  {name}");
            }
            catch (SkipException ex)
            {
                _skipped++;
                Console.WriteLine($"  SKIP  {name}: {ex.Message}");
            }
            catch (Exception ex)
            {
                _failed++;
                Console.WriteLine($"  FAIL  {name}: {ex.Message}");
            }
        }

        private sealed class SkipException : Exception
        {
            public SkipException(string message) : base(message) { }
        }

        private static void Assert(bool cond, string msg)
        {
            if (!cond) throw new Exception(msg);
        }

        private static void AssertEq<T>(T expected, T actual, string msg)
        {
            if (!EqualityComparer<T>.Default.Equals(expected, actual))
                throw new Exception($"{msg}: 期望 [{expected}] 实际 [{actual}]");
        }

        private static void AssertSeq<T>(IEnumerable<T> expected, IEnumerable<T> actual, string msg)
        {
            var e = expected.ToList();
            var a = actual.ToList();
            if (e.Count != a.Count || !e.SequenceEqual(a))
                throw new Exception($"{msg}: 期望 [{string.Join(",", e)}] 实际 [{string.Join(",", a)}]");
        }

        private static string NewTempDir()
        {
            var dir = Path.Combine(Path.GetTempPath(), "wxd_reg_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            return dir;
        }

        // git 松散对象是只读文件，直接 Directory.Delete 会抛 UnauthorizedAccess；先清只读属性再删
        private static void DeleteTempDir(string dir)
        {
            if (!Directory.Exists(dir)) return;
            foreach (var f in Directory.EnumerateFiles(dir, "*", SearchOption.AllDirectories))
            {
                try { File.SetAttributes(f, FileAttributes.Normal); } catch { }
            }
            try { Directory.Delete(dir, true); } catch { }
        }

        private static string MakeFile(string root, params string[] parts)
        {
            var dir = root;
            for (int i = 0; i < parts.Length - 1; i++)
            {
                dir = Path.Combine(dir, parts[i]);
                Directory.CreateDirectory(dir);
            }
            var file = Path.Combine(dir, parts[^1]);
            File.WriteAllText(file, "x");
            return file;
        }

        private static FileTableData MakeFd(params (int origIdx, string[] cells)[] rows)
        {
            var fd = new FileTableData();
            foreach (var (origIdx, cells) in rows)
            {
                var r = new TableRow { OriginalRowIndex = origIdx };
                r.Cells.AddRange(cells);
                fd.Rows.Add(r);
            }
            return fd;
        }

        private static bool GitAvailable()
        {
            try
            {
                var psi = new ProcessStartInfo("git", "--version")
                {
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                };
                using var p = Process.Start(psi);
                if (p == null) return false;
                p.WaitForExit(10_000);
                return p.HasExited && p.ExitCode == 0;
            }
            catch { return false; }
        }

        private static void Git(string repo, string args)
        {
            var psi = new ProcessStartInfo("git", args)
            {
                WorkingDirectory = repo,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
            };
            using var p = Process.Start(psi)!;
            var err = p.StandardError.ReadToEnd();
            p.StandardOutput.ReadToEnd();
            p.WaitForExit(30_000);
            if (p.ExitCode != 0)
                throw new Exception($"git {args} 失败 (exit={p.ExitCode}): {err.Trim()}");
        }
    }
}