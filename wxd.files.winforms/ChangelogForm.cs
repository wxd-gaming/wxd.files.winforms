using System;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>更新记录展示窗体，显示最近 N 次 Release 的更新日志</summary>
    /// <remarks>排版和帮助窗完全一样，只是内容改成异步拉取，所以直接继承，别再抄一份配色表。</remarks>
    public class ChangelogForm : HelpForm
    {
        public ChangelogForm() : base("更新记录", "正在获取更新记录...")
        {
        }

        protected override async void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            string text;
            try
            {
                var notes = await UpdateHelper.GetRecentReleaseNotesAsync(5);
                text = string.IsNullOrWhiteSpace(notes) ? "暂无更新记录。" : notes;
            }
            catch
            {
                text = "获取更新记录失败，请检查网络连接后重试。";
            }

            _rtbContent.Text = text;
            _rtbContent.SelectionStart = 0;
            _rtbContent.ScrollToCaret();
        }

        /// <summary>便捷方法：显示更新记录窗体</summary>
        public static void ShowChangelog(IWin32Window owner)
        {
            using var form = new ChangelogForm();
            form.ShowDialog(owner);
        }
    }
}
