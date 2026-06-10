using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>独立表格查看窗体（用于关联打开，不包含左侧文件列表和目录树）</summary>
    public class CsvReaderStandaloneForm : Form
    {
        /// <summary>使用配置对象构造（推荐）</summary>
        public CsvReaderStandaloneForm(TableBrowserOptions options, LinkOpenOptions? linkOptions = null)
        {
            if (options == null) throw new ArgumentNullException(nameof(options));

            var linkOpt = linkOptions ?? new LinkOpenOptions();
            Text = string.IsNullOrEmpty(linkOpt.DisplayName) ? Path.GetFileName(options.FilePath) : linkOpt.DisplayName;
            if (linkOpt.IsLinkedFromCell)
            {
                ClientSize = new Size(1300, 350);
                MinimumSize = new Size(300, 200);
            }
            else
            {
                ClientSize = new Size(1060, 620);
                MinimumSize = new Size(560, 350);
            }
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Microsoft YaHei UI", 9f);
            try { Icon = Icon.ExtractAssociatedIcon(options.FilePath); } catch { }
            BackColor = Color.FromArgb(240, 242, 245);

            var viewer = new TableBrowserControl(options) { Dock = DockStyle.Fill };
            viewer.SetHighlight(linkOpt.HighlightColumnIndex, linkOpt.HighlightValue, linkOpt.HighlightColumnName, linkOpt.LinkConfig);
            Controls.Add(viewer);
        }

        /// <summary>旧构造函数（已废弃，委托给新构造函数）</summary>
        [Obsolete("请使用 CsvReaderStandaloneForm(TableBrowserOptions, LinkOpenOptions?) 代替")]
        public CsvReaderStandaloneForm(string filePath, int highlightColumnIndex = -1, string highlightValue = "", int frozenRows = -1, string highlightColumnName = "", int headerRowIndex = -1, bool isLinkedFromCell = false, ColumnLinkConfig? linkConfig = null, string displayName = "")
            : this(
                new TableBrowserOptions(filePath)
                {
                    RootDir = Path.GetDirectoryName(filePath) ?? "",
                    FrozenRows = frozenRows,
                    HeaderRowIndex = headerRowIndex,
                },
                new LinkOpenOptions
                {
                    HighlightColumnIndex = highlightColumnIndex,
                    HighlightValue = highlightValue,
                    HighlightColumnName = highlightColumnName,
                    LinkConfig = linkConfig,
                    DisplayName = displayName,
                    IsLinkedFromCell = isLinkedFromCell,
                })
        { }
    }
}
