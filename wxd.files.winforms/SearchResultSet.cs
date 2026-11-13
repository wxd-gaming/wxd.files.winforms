using System;
using System.Collections.Generic;
using System.Threading;

namespace wxd.files.winforms
{
    /// <summary>
    /// 搜索结果集：匹配行号按升序分段存放。
    ///
    /// 为什么不继续用 List&lt;long&gt;：并行搜索各块的完成顺序是乱的，把一个块的行号
    /// 逐个 Insert 进同一个有序 List，每次都是 O(n) 内存搬移；千万级匹配时光合并
    /// 就要在 UI 线程跑几十分钟，界面直接假死。而块号本身就是行号顺序（块按字节
    /// 均分、字节有序即行号有序），所以只要让每块按块号落位、彼此不合并，追加就
    /// 退化成一次 O(段数) 发布，全局二分照样成立。
    ///
    /// 线程模型：写侧（后台搜索线程）在 _gate 内落位并重建只读视图；读侧（UI 线程
    /// 导航、导出线程遍历）只做一次 Volatile.Read 拿到一致快照，段数组发布后不再
    /// 改动，因此读取全程无锁也不会读到撕裂状态。
    /// </summary>
    public sealed class SearchResultSet
    {
        /// <summary>发布后不可变的读侧快照</summary>
        private sealed class View
        {
            public static readonly View Empty = new(Array.Empty<long[]>(), new long[] { 0 }, 0L);

            public readonly long[][] Segments;  // 按行号升序，已剔除空段
            public readonly long[] Starts;      // Starts[i] = 前 i 段元素总数，长度 = Segments.Length + 1
            public readonly long Total;

            public View(long[][] segments, long[] starts, long total)
            {
                Segments = segments;
                Starts = starts;
                Total = total;
            }
        }

        private readonly object _gate = new();
        private readonly List<(int Slot, long[] Lines)> _slots = new();
        private View _view = View.Empty;
        private bool _sealed;

        /// <summary>已收集的匹配总数（无锁读，可能是稍旧的一致快照）</summary>
        public long Count => Volatile.Read(ref _view).Total;

        /// <summary>
        /// 扫描是否已正常跑完。未封口的容器只有"到目前为止"的匹配：搜索还在跑，
        /// 或中途被取消/切走关键词留下了半截结果。导出必须等封口，否则产物残缺。
        /// </summary>
        public bool IsSealed => Volatile.Read(ref _sealed);

        /// <summary>扫描正常结束时封口；被取消的扫描不要调用</summary>
        public void Seal() => Volatile.Write(ref _sealed, true);

        /// <summary>估算占用字节数：行号 8 字节 + 每段数组头</summary>
        public long EstimatedBytes
        {
            get
            {
                var v = Volatile.Read(ref _view);
                return v.Total * 8L + v.Segments.Length * 64L + 128L;
            }
        }

        /// <summary>
        /// 发布一段行号。slot 决定段与段之间的先后（并行搜索传块号，顺序扫描传递增号），
        /// 同 slot 重复发布以最后一次为准，便于块重扫。
        /// </summary>
        public void Publish(int slot, long[] lines)
        {
            if (lines == null) lines = Array.Empty<long>();
            lock (_gate)
            {
                int at = 0;
                while (at < _slots.Count && _slots[at].Slot < slot) at++;
                if (at < _slots.Count && _slots[at].Slot == slot)
                    _slots[at] = (slot, lines);
                else
                    _slots.Insert(at, (slot, lines));
                RebuildViewLocked();
            }
        }

        /// <summary>追加一段，槽位取当前最大槽位 + 1（顺序扫描路径用）</summary>
        public void AppendSegment(long[] lines)
        {
            lock (_gate)
            {
                int slot = _slots.Count == 0 ? 0 : _slots[^1].Slot + 1;
                _slots.Add((slot, lines ?? Array.Empty<long>()));
                RebuildViewLocked();
            }
        }

        /// <summary>取当前段快照（引用共享，段发布后不再修改），供顺序遍历与写盘使用</summary>
        public long[][] SnapshotSegments() => Volatile.Read(ref _view).Segments;

        public long this[long index]
        {
            get
            {
                var v = Volatile.Read(ref _view);
                if (index < 0 || index >= v.Total) throw new IndexOutOfRangeException($"搜索结果索引 {index} 超出范围 {v.Total}");
                return GetAt(v, index);
            }
        }

        /// <summary>
        /// 二分定位：等于 line 时返回其下标，否则返回最后一个小于 line 的下标（全大于则 0）。
        /// 空集返回 -1。语义与旧的 FindNearestIndex(List&lt;long&gt;) 一致。
        /// </summary>
        public long IndexOfNearest(long line)
        {
            var v = Volatile.Read(ref _view);
            if (v.Total == 0) return -1;
            long lo = 0, hi = v.Total - 1;
            while (lo <= hi)
            {
                long mid = lo + ((hi - lo) >> 1);
                long val = GetAt(v, mid);
                if (val == line) return mid;
                if (val < line) lo = mid + 1;
                else hi = mid - 1;
            }
            return hi >= 0 ? hi : 0;
        }

        private static long GetAt(View v, long index)
        {
            // 找最大的 k 使 Starts[k] <= index
            int lo = 0, hi = v.Segments.Length - 1, k = 0;
            while (lo <= hi)
            {
                int mid = (lo + hi) / 2;
                if (v.Starts[mid] <= index) { k = mid; lo = mid + 1; }
                else hi = mid - 1;
            }
            return v.Segments[k][index - v.Starts[k]];
        }

        private void RebuildViewLocked()
        {
            int n = 0;
            long total = 0;
            foreach (var (_, lines) in _slots)
            {
                if (lines.Length == 0) continue;
                n++;
                total += lines.Length;
            }
            if (n == 0)
            {
                Volatile.Write(ref _view, View.Empty);
                return;
            }
            var segs = new long[n][];
            int i = 0;
            foreach (var (_, lines) in _slots)
            {
                if (lines.Length == 0) continue;
                segs[i++] = lines;
            }
            var starts = new long[n + 1];
            for (int s = 0; s < n; s++) starts[s + 1] = starts[s] + segs[s].Length;
            Volatile.Write(ref _view, new View(segs, starts, total));
        }
    }
}
