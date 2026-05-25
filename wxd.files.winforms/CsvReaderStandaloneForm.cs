using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>独立表格查看窗体（用于关联打开，不包含左侧文件列表和目录树）</summary>
    public class CsvReaderStandaloneForm : Form
    {
        public CsvReaderStandaloneForm(string filePath)
        {
            var fileName = Path.GetFileName(filePath);
            Text = fileName;
            ClientSize = new Size(1200, 800);
            MinimumSize = new Size(600, 400);
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Microsoft YaHei UI", 9f);
            try { Icon = Icon.ExtractAssociatedIcon(filePath); } catch { }
            BackColor = Color.FromArgb(240, 242, 245);

            var ext = Path.GetExtension(filePath).ToLowerInvariant();
            var dir = Path.GetDirectoryName(filePath) ?? "";
            Control viewer;
            if (ext == ".csv")
                viewer = new CsvViewerControl(filePath, rootDir: dir) { Dock = DockStyle.Fill };
            else
                viewer = new ExcelViewerControl(filePath, rootDir: dir) { Dock = DockStyle.Fill };

            Controls.Add(viewer);
        }
    }
}
