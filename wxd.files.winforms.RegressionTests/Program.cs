using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json.Nodes;
using System.Windows.Forms;
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

            // 性能重构：索引映射 / 过滤词预编译 / 排序键预计算
            Run("映射 反查表基本语义", Test_BuildReverseMap);
            Run("映射 冻结行与内容行分组", Test_MapMatchCellsFrozenAndBody);
            Run("映射 被过滤掉的行跳过", Test_MapMatchCellsSkipsHiddenRows);
            Run("映射 无冻结行按原行号直映", Test_MapMatchCellsDirect);
            Run("过滤词 单串/OR/AND", Test_ContentFilterMatcher);
            Run("过滤词 边界与大小写", Test_ContentFilterMatcherEdges);
            Run("排序 数值升序/降序", Test_SortRowIndicesNumeric);
            Run("排序 同键稳定/无效列", Test_SortRowIndicesStable);
            Run("行数据 子集继承物理行号", Test_BuildFileDataSubsetInherit);
            Run("CSV 写出转义/无BOM", Test_CsvWriteRoundTrip);

            // 文本比对引擎
            Run("文本比对 相同/新增/删除/修改", Test_TextDiffBasic);
            Run("文本比对 多段差异可还原", Test_TextDiffMultiHunk);
            Run("文本比对 空侧与只保留差异行", Test_TextDiffEmptyAndFilter);
            Run("文本比对 差异过大退化不崩", Test_TextDiffFallback);

            // 表格比对（csv / excel）
            Run("表格比对 相同/单元格差异/增删行", Test_TableDiffBasic);
            Run("表格比对 列数不齐/只保留差异行", Test_TableDiffRaggedAndOnlyDifferences);
            Run("表格读取 CSV 引号/换行/分隔符", Test_SimpleTableCsv);
            Run("表格读取 CSV 编码 GBK", Test_SimpleTableCsvGbk);
            Run("表格读取 Excel 往返", Test_SimpleTableExcelRoundTrip);

            // JSON 面板
            Run("JSON 格式化/压缩/中文不转义", Test_JsonFormatAndMinify);
            Run("JSON 递归 key 排序", Test_JsonSortKeys);
            Run("JSON 解析错误与宽松语法", Test_JsonParseErrors);
            Run("JSON 节点计数与顶层 null", Test_JsonCountAndNull);

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

        // ==================== 性能重构（索引映射 / 过滤词 / 排序键） ====================

        private static void Test_BuildReverseMap()
        {
            var map = TableIndexMapper.BuildReverseMap(new[] { 5, 2, 9 }, 10);
            AssertEq(0, map[5], "5 → 显示行 0");
            AssertEq(1, map[2], "2 → 显示行 1");
            AssertEq(2, map[9], "9 → 显示行 2");
            AssertEq(-1, map[0], "未出现的索引为 -1");

            var empty = TableIndexMapper.BuildReverseMap(null, 3);
            AssertEq(-1, empty[0], "null 列表应全部为 -1");

            var outOfRange = TableIndexMapper.BuildReverseMap(new[] { 7 }, 2);
            AssertEq(-1, outOfRange[1], "越界的原始索引应被忽略而不是抛异常");
        }

        private static void Test_MapMatchCellsFrozenAndBody()
        {
            // 显示列：原始 0 → 显示 0，原始 2 → 显示 1（原始列 1 被隐藏）
            var visibleCols = new List<int> { 0, 2 };
            // 内容行显示顺序：原始 10 → 显示 0，原始 12 → 显示 1
            var visibleRows = new List<int> { 10, 12 };
            var merged = new Dictionary<int, HashSet<int>>
            {
                [1] = new HashSet<int> { 0, 2 },    // 冻结行 1
                [10] = new HashSet<int> { 2 },      // 内容行 → 显示 0
                [12] = new HashSet<int> { 0, 2 },   // 内容行 → 显示 1
            };

            var (header, body) = TableIndexMapper.MapMatchCellsToDisplay(
                merged, visibleCols, visibleRows, frozenRows: 2, columnCount: 3, rowCount: 20);

            AssertEq(2, header[1].Count, "冻结行应保留两列匹配");
            Assert(header[1].Contains(0) && header[1].Contains(1), "原始列 0/2 应映射为显示列 0/1");

            AssertEq(2, body.Count, "两行内容行应各有映射");
            AssertEq(1, body[0].Count, "内容行 10 只命中原始列 2");
            Assert(body[0].Contains(1), "原始列 2 的显示索引应为 1");
            AssertEq(2, body[1].Count, "内容行 12 两列都命中");
        }

        private static void Test_MapMatchCellsSkipsHiddenRows()
        {
            var merged = new Dictionary<int, HashSet<int>> { [7] = new HashSet<int> { 0 } };
            var (header, body) = TableIndexMapper.MapMatchCellsToDisplay(
                merged, new List<int> { 0 }, new List<int> { 1, 2 },
                frozenRows: 0, columnCount: 1, rowCount: 10);

            AssertEq(0, header.Count, "无冻结行时不应产生表头映射");
            AssertEq(0, body.Count, "被过滤掉的行不应出现在结果里");
        }

        private static void Test_MapMatchCellsDirect()
        {
            var merged = new Dictionary<int, HashSet<int>> { [3] = new HashSet<int> { 1 } };
            var (_, body) = TableIndexMapper.MapMatchCellsToDisplay(
                merged, new List<int> { 0, 1 }, null,
                frozenRows: 0, columnCount: 2, rowCount: 5);

            AssertEq(1, body[3].Single(), "无过滤无排序时显示行 = 原始行");
        }

        private static void Test_ContentFilterMatcher()
        {
            Assert(ContentFilterMatcher.Parse("abc").Matches("xxABCxx"), "单串应大小写不敏感");
            Assert(!ContentFilterMatcher.Parse("abc").Matches("xxabdxx"), "单串不命中应返回 false");

            var or = ContentFilterMatcher.Parse("abc||def");
            Assert(or.Matches("DEF"), "OR：命中第二个词");
            Assert(or.Matches("xabcdef"), "OR：命中第一个词");
            Assert(!or.Matches("xyz"), "OR：都不命中");

            var and = ContentFilterMatcher.Parse("abc&&def");
            Assert(and.Matches("abcdef"), "AND：两个词都在");
            Assert(!and.Matches("abc"), "AND：缺一个词不命中");
        }

        private static void Test_ContentFilterMatcherEdges()
        {
            Assert(ContentFilterMatcher.Parse("").Matches("任意"), "空词表 = 全部命中");
            Assert(!ContentFilterMatcher.Parse("|||").Matches("abc"), "只有分隔符时 OR 不命中");
            Assert(ContentFilterMatcher.Parse("&&&").Matches("abc"), "只有分隔符时 AND 视为全部通过");
            Assert(!ContentFilterMatcher.Parse("abc").Matches(""), "空单元格不命中非空过滤词");
        }

        private static void Test_SortRowIndicesNumeric()
        {
            var dt = new DataTable();
            dt.Columns.Add("v", typeof(string));
            foreach (var v in new[] { "10", "2", "1" }) dt.Rows.Add(v);

            var asc = TableBrowserControl.SortRowIndices(dt, new List<int> { 0, 1, 2 }, 0, SortOrder.Ascending);
            AssertSeq(new[] { 2, 1, 0 }, asc, "数值列应按数值升序（1,2,10），不是字典序");

            var desc = TableBrowserControl.SortRowIndices(dt, new List<int> { 0, 1, 2 }, 0, SortOrder.Descending);
            AssertSeq(new[] { 0, 1, 2 }, desc, "数值列降序");
        }

        private static void Test_SortRowIndicesStable()
        {
            var dt = new DataTable();
            dt.Columns.Add("v", typeof(string));
            foreach (var v in new[] { "b", "B", "a" }) dt.Rows.Add(v);

            var sorted = TableBrowserControl.SortRowIndices(dt, new List<int> { 0, 1, 2 }, 0, SortOrder.Ascending);
            AssertSeq(new[] { 2, 0, 1 }, sorted, "同键按输入顺序稳定排列（b 先于 B）");

            var invalidColumn = TableBrowserControl.SortRowIndices(dt, new List<int> { 0, 1 }, 9, SortOrder.Ascending);
            AssertSeq(new[] { 0, 1 }, invalidColumn, "无效排序列应原样返回");
        }

        private static void Test_BuildFileDataSubsetInherit()
        {
            var source = new FileTableData { FilePath = "a.csv", Type = TableFileType.Csv };
            source.ColumnNames.Add("v");
            source.Rows.Add(new TableRow { OriginalRowIndex = 0, PhysicalStartLine = 1, PhysicalEndLine = 1, Cells = new List<string> { "a" } });
            source.Rows.Add(new TableRow { OriginalRowIndex = 1, PhysicalStartLine = 3, PhysicalEndLine = 4, Cells = new List<string> { "b" } });
            source.Rows.Add(new TableRow { OriginalRowIndex = 2, PhysicalStartLine = 5, PhysicalEndLine = 5, Cells = new List<string> { "c" } });

            var subset = new DataTable();
            subset.Columns.Add("v", typeof(string));
            subset.Rows.Add("x");
            subset.Rows.Add("y");
            subset.Rows.Add("z");

            var built = TableBrowserControl.BuildFileData(subset, source);
            AssertEq(3, built.Rows.Count, "子集行数");
            AssertSeq(new[] { 1, 3, 5 }, built.Rows.Select(r => r.PhysicalStartLine), "子集应继承原行的物理起始行号");
            AssertSeq(new[] { 0, 1, 2 }, built.Rows.Select(r => r.OriginalRowIndex), "OriginalRowIndex 应指向源行号");
        }

        private static void Test_CsvWriteRoundTrip()
        {
            var dir = NewTempDir();
            try
            {
                var dt = new DataTable();
                dt.Columns.Add("a", typeof(string));
                dt.Columns.Add("b", typeof(string));
                dt.Rows.Add("x,y", "line1\nline2");
                dt.Rows.Add("quote\"d", "");

                var path = Path.Combine(dir, "out.csv");
                FileDataCache.WriteCsvRows(path, dt, ',');

                var bytes = File.ReadAllBytes(path);
                Assert(!(bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF),
                    "CSV 不应写 UTF-8 BOM");
                var text = File.ReadAllText(path);
                AssertEq("\"x,y\",\"line1\nline2\"\r\n\"quote\"\"d\",\r\n", text,
                    "CSV 转义与换行格式（含引号、内嵌换行）");
            }
            finally
            {
                DeleteTempDir(dir);
            }
        }

        // ==================== 文本比对引擎 ====================

        private static void AssertDiffReconstruct(string[] left, string[] right, List<TextDiffLine> diff)
        {
            var rebuiltLeft = diff.Where(d => d.LeftLineNumber != null).Select(d => d.LeftText).ToArray();
            var rebuiltRight = diff.Where(d => d.RightLineNumber != null).Select(d => d.RightText).ToArray();
            AssertSeq(left, rebuiltLeft, "按左侧行号拼回去必须还原左文件");
            AssertSeq(right, rebuiltRight, "按右侧行号拼回去必须还原右文件");

            int lastLeft = 0, lastRight = 0;
            foreach (var line in diff)
            {
                if (line.LeftLineNumber != null)
                {
                    Assert(line.LeftLineNumber.Value == lastLeft + 1, "左侧行号必须连续递增");
                    lastLeft = line.LeftLineNumber.Value;
                }
                if (line.RightLineNumber != null)
                {
                    Assert(line.RightLineNumber.Value == lastRight + 1, "右侧行号必须连续递增");
                    lastRight = line.RightLineNumber.Value;
                }
                Assert(!(line.LeftLineNumber == null && line.RightLineNumber == null), "不允许两侧都没有行号的对齐行");
            }
        }

        private static void Test_TextDiffBasic()
        {
            var same = new[] { "a", "b", "c" };
            var sameDiff = TextDiffEngine.Compare(same, same);
            AssertEq(3, sameDiff.Count, "相同文件行数不变");
            Assert(sameDiff.All(d => d.Kind == TextDiffKind.Same), "相同文件不应有差异");
            AssertDiffReconstruct(same, same, sameDiff);

            var left = new[] { "a", "b", "c" };
            var right = new[] { "a", "x", "b", "c" };
            var inserted = TextDiffEngine.Compare(left, right);
            AssertEq(1, inserted.Count(d => d.Kind == TextDiffKind.Added), "插入一行");
            AssertDiffReconstruct(left, right, inserted);

            var deleted = TextDiffEngine.Compare(right, left);
            AssertEq(1, deleted.Count(d => d.Kind == TextDiffKind.Deleted), "删除一行");
            AssertDiffReconstruct(right, left, deleted);

            var modifiedLeft = new[] { "a", "b", "c" };
            var modifiedRight = new[] { "a", "B", "c" };
            var modified = TextDiffEngine.Compare(modifiedLeft, modifiedRight);
            AssertEq(1, modified.Count(d => d.Kind == TextDiffKind.Modified), "修改一行");
            AssertEq(0, modified.Count(d => d.Kind == TextDiffKind.Added || d.Kind == TextDiffKind.Deleted),
                "一对一替换不应再拆成新增/删除");
            AssertDiffReconstruct(modifiedLeft, modifiedRight, modified);
        }

        private static void Test_TextDiffMultiHunk()
        {
            var left = new List<string>();
            var right = new List<string>();
            for (int i = 0; i < 200; i++)
            {
                left.Add("line" + i);
                right.Add(i == 5 || i == 120 ? "changed" + i : "line" + i);
            }
            right.Insert(60, "extra");
            right.RemoveAt(150);

            var diff = TextDiffEngine.Compare(left, right);
            Assert(diff.Count(d => d.Kind == TextDiffKind.Modified) >= 2, "两处修改");
            Assert(diff.Count(d => d.Kind == TextDiffKind.Added) >= 1, "一处新增");
            Assert(diff.Count(d => d.Kind == TextDiffKind.Deleted) >= 1, "一处删除");
            AssertDiffReconstruct(left.ToArray(), right.ToArray(), diff);
        }

        private static void Test_TextDiffEmptyAndFilter()
        {
            var emptyDiff = TextDiffEngine.Compare(Array.Empty<string>(), new[] { "a", "b" });
            AssertEq(2, emptyDiff.Count(d => d.Kind == TextDiffKind.Added), "左空 → 全是新增");
            AssertDiffReconstruct(Array.Empty<string>(), new[] { "a", "b" }, emptyDiff);

            var rightEmpty = TextDiffEngine.Compare(new[] { "a" }, Array.Empty<string>());
            AssertEq(1, rightEmpty.Count(d => d.Kind == TextDiffKind.Deleted), "右空 → 全是删除");
            AssertDiffReconstruct(new[] { "a" }, Array.Empty<string>(), rightEmpty);

            var diff = TextDiffEngine.Compare(new[] { "a", "b" }, new[] { "a", "x" });
            var only = TextDiffEngine.OnlyDifferences(diff);
            AssertEq(1, only.Count, "只保留差异行应过滤掉相同行");
            AssertEq(TextDiffKind.Modified, only[0].Kind, "过滤后保留的是修改行");
        }

        private static void Test_TextDiffFallback()
        {
            // 两侧完全不同、行数超过 MaxEditDistance，走"整块替换"退化路径
            var left = Enumerable.Range(0, 3000).Select(i => "L" + i).ToArray();
            var right = Enumerable.Range(0, 3000).Select(i => "R" + i).ToArray();

            var diff = TextDiffEngine.Compare(left, right);
            Assert(diff.Count >= 3000, "退化后仍有全部行");
            Assert(diff.All(d => d.Kind == TextDiffKind.Modified), "等长完全不同的文件应两两配成修改行");
            AssertDiffReconstruct(left, right, diff);
        }

        // ==================== 表格比对 ====================

        private static void Test_TableDiffBasic()
        {
            var left = new List<string[]>
            {
                new[] { "id", "name", "hp" },
                new[] { "1", "slime", "10" },
                new[] { "2", "goblin", "20" }
            };
            var right = new List<string[]>
            {
                new[] { "id", "name", "hp" },
                new[] { "1", "slime", "12" },
                new[] { "2", "goblin", "20" },
                new[] { "3", "boss", "100" }
            };

            var diff = TableDiffEngine.Compare(left, right);
            AssertEq(4, diff.Count, "按行数多的一侧对齐");
            AssertEq(TableRowDiffKind.Same, diff[0].Kind, "表头行相同");
            AssertEq(TableRowDiffKind.Modified, diff[1].Kind, "第 2 行有一格变化");
            Assert(diff[1].DiffColumns.SetEquals(new[] { 2 }), "差异列应为 C 列（索引 2）");
            AssertEq(TableRowDiffKind.Same, diff[2].Kind, "第 3 行相同");
            AssertEq(TableRowDiffKind.Added, diff[3].Kind, "右表多出一行");
            AssertEq(4, diff[3].RightRowNumber, "新增行右侧行号（第 4 行）");

            var reverse = TableDiffEngine.Compare(right, left);
            AssertEq(TableRowDiffKind.Deleted, reverse[3].Kind, "反向比对最后一行是删除");
            AssertEq(0, TableDiffEngine.Compare(left, left).Count(d => d.Kind != TableRowDiffKind.Same),
                "同一个表比对不应有差异");
        }

        private static void Test_TableDiffRaggedAndOnlyDifferences()
        {
            var left = new List<string[]> { new[] { "a" } };
            var right = new List<string[]> { new[] { "a", "b" } };
            var diff = TableDiffEngine.Compare(left, right);
            AssertEq(TableRowDiffKind.Modified, diff[0].Kind, "列数不同也算修改");
            Assert(diff[0].DiffColumns.Contains(1), "多出来的列应标为差异");

            var only = TableDiffEngine.OnlyDifferences(TableDiffEngine.Compare(
                new List<string[]> { new[] { "a" }, new[] { "b" } },
                new List<string[]> { new[] { "a" }, new[] { "x" } }));
            AssertEq(1, only.Count, "只保留差异行应过滤掉相同行");
            AssertEq(TableRowDiffKind.Modified, only[0].Kind, "保留的是修改行");
        }

        private static void Test_SimpleTableCsv()
        {
            var dir = NewTempDir();
            try
            {
                var path = Path.Combine(dir, "t.csv");
                File.WriteAllText(path, "id,name,note\n1,\"a,b\",x\n2,\"line1\nline2\",y\n", Encoding.UTF8);

                var rows = SimpleTableReader.ReadRows(path);
                AssertEq(3, rows.Count, "三条记录");
                AssertSeq(new[] { "id", "name", "note" }, rows[0], "第一行原样当数据");
                AssertSeq(new[] { "1", "a,b", "x" }, rows[1], "引号内的逗号不切列");
                AssertEq("line1\nline2", rows[2][1], "引号内的换行不切行");

                var tabbed = Path.Combine(dir, "t-tab.csv");
                File.WriteAllText(tabbed, "a\tb\tc\r\n1\t2\t3\r\n", Encoding.UTF8);
                var tabRows = SimpleTableReader.ReadRows(tabbed);
                AssertEq('\t', SimpleTableReader.DetectCsvDelimiter("a\tb\tc\r\n1\t2\t3"), "分隔符应识别为制表符");
                AssertEq(3, tabRows[1].Length, "制表符分隔应切出三列");
            }
            finally
            {
                DeleteTempDir(dir);
            }
        }

        private static void Test_SimpleTableCsvGbk()
        {
            var dir = NewTempDir();
            try
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                var gbk = Encoding.GetEncoding(936);
                var path = Path.Combine(dir, "gbk.csv");
                File.WriteAllText(path, "编号,名称\r\n1,史莱姆\r\n", gbk);

                var rows = SimpleTableReader.ReadRows(path);
                AssertEq("编号", rows[0][0], "GBK 中文列名");
                AssertEq("史莱姆", rows[1][1], "GBK 中文内容");
            }
            finally
            {
                DeleteTempDir(dir);
            }
        }

        private static void Test_SimpleTableExcelRoundTrip()
        {
            var dir = NewTempDir();
            try
            {
                var path = Path.Combine(dir, "t.xlsx");
                var table = new DataTable();
                table.Columns.Add("c0", typeof(string));
                table.Columns.Add("c1", typeof(string));
                table.Rows.Add("id", "name");
                table.Rows.Add("1", "slime");
                MiniExcelLibs.MiniExcel.SaveAs(path, table, printHeader: false);

                Assert(SimpleTableReader.IsTableFile(path), "xlsx 应识别为表格");
                var rows = SimpleTableReader.ReadRows(path);
                AssertEq(2, rows.Count, "Excel 两行");
                AssertEq("id", rows[0][0], "Excel 单元格内容");
                AssertEq("slime", rows[1][1], "Excel 第二行内容");
            }
            finally
            {
                DeleteTempDir(dir);
            }
        }

        // ==================== JSON 面板 ====================

        private static void Test_JsonFormatAndMinify()
        {
            var result = JsonToolEngine.Parse("""{"b":1,"a":{"d":2,"c":3},"name":"史莱姆"}""");
            Assert(result.Success, "合法 JSON 应解析成功");

            var formatted = JsonToolEngine.Format(result.Root);
            Assert(formatted.Contains("\n  \"b\""), "应使用 2 空格缩进");
            Assert(formatted.Contains("史莱姆"), "中文不应被转义");

            var minified = JsonToolEngine.Minify(result.Root);
            Assert(!minified.Contains('\n'), "压缩后不应有换行");
            Assert(!minified.Contains(' '), "压缩后不应有空格");
            Assert(minified.Contains("\"name\":\"史莱姆\""), "压缩后结构完整");
        }

        private static void Test_JsonSortKeys()
        {
            var result = JsonToolEngine.Parse("""{"b":{"d":1,"c":2},"a":3,"arr":[{"z":1,"y":2}]}""");
            Assert(result.Success, "解析成功");

            var formatted = JsonToolEngine.Format(result.Root, sortKeys: true);
            var expected = """
{
  "a": 3,
  "arr": [
    {
      "y": 2,
      "z": 1
    }
  ],
  "b": {
    "c": 2,
    "d": 1
  }
}
""";
            AssertEq(expected.ReplaceLineEndings("\n"), formatted.ReplaceLineEndings("\n"),
                "递归排序：对象 key 升序、数组顺序不变");
        }

        private static void Test_JsonParseErrors()
        {
            var bad = JsonToolEngine.Parse("""{"a": }""");
            Assert(!bad.Success, "非法 JSON 应失败");
            Assert(bad.Error != null && bad.Error.Contains("语法错误"), "错误信息应可读");

            // 尾逗号 + 注释：工具场景里很常见，按宽松模式接受
            var lenient = JsonToolEngine.Parse("""
{
  "a": 1, // 注释
  "b": 2,
}
""");
            Assert(lenient.Success, "允许注释与尾逗号");
            Assert(lenient.Root is JsonObject obj && obj.Count == 2, "两个字段");
        }

        private static void Test_JsonCountAndNull()
        {
            var result = JsonToolEngine.Parse("""{"a":1,"b":[2,3]}""");
            AssertEq(5, JsonToolEngine.CountNodes(result.Root), "对象1 + 值1 + 数组1 + 两项2");

            var nullResult = JsonToolEngine.Parse("null");
            Assert(nullResult.Success && nullResult.Root == null, "顶层 null 也是合法 JSON");
            AssertEq("null", JsonToolEngine.Format(nullResult.Root), "null 原样输出");
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