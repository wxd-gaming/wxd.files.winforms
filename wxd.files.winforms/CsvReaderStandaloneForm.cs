using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>独立表格查看窗体（用于关联打开，不包含左侧文件列表和目录树）</summary>
    public class CsvReaderStandaloneForm : Form
    {
        public CsvReaderStandaloneForm(string filePath, int highlightColumnIndex = -1, string highlightValue = "", int frozenRows = -1, string highlightColumnName = "", int headerRowIndex = -1, bool isLinkedFromCell = false, ColumnLinkConfig? linkConfig = null, string displayName = "")
        {
            Text = string.IsNullOrEmpty(displayName) ? Path.GetFileName(filePath) : displayName;
            if (isLinkedFromCell)
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
            try { Icon = Icon.ExtractAssociatedIcon(filePath); } catch { }
            BackColor = Color.FromArgb(240, 242, 245);

            var dir = Path.GetDirectoryName(filePath) ?? "";
            var viewer = new TableBrowserControl(filePath, rootDir: dir, frozenRows: frozenRows, headerRowIndex: headerRowIndex) { Dock = DockStyle.Fill };
            viewer.SetHighlight(highlightColumnIndex, highlightValue, highlightColumnName, linkConfig);
            Controls.Add(viewer);
        }
    }
}
