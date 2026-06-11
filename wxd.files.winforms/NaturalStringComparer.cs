namespace wxd.files.winforms
{
    /// <summary>自然字符串比较器，支持数字排序</summary>
    public class NaturalStringComparer : IComparer<string?>
    {
        private readonly SortOrder _sortOrder;

        public NaturalStringComparer(SortOrder sortOrder)
        {
            _sortOrder = sortOrder;
        }

        public int Compare(string? x, string? y)
        {
            if (x == null && y == null) return 0;
            if (x == null) return _sortOrder == SortOrder.Ascending ? -1 : 1;
            if (y == null) return _sortOrder == SortOrder.Ascending ? 1 : -1;

            // 尝试数字比较
            if (double.TryParse(x, out double numX) && double.TryParse(y, out double numY))
            {
                var result = numX.CompareTo(numY);
                return _sortOrder == SortOrder.Ascending ? result : -result;
            }

            // 字符串比较
            var strResult = string.Compare(x, y, StringComparison.OrdinalIgnoreCase);
            return _sortOrder == SortOrder.Ascending ? strResult : -strResult;
        }
    }
}
