using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>
    /// .NET 10 的 SplitContainer 在显示会话切换时会从 RepaintSplitterRect 抛 GDI+ ExternalException，
    /// 该异常跨出窗口过程会直接杀进程；.NET 11 才在控件内部兜住。
    /// </summary>
    internal sealed class SafeSplitContainer : SplitContainer
    {
        private bool _retryQueued;
        private int _retryCount;

        protected override void OnLayout(LayoutEventArgs e)
        {
            try
            {
                base.OnLayout(e);
                _retryCount = 0;
            }
            catch (ExternalException)
            {
                // 异常会打断本次布局，并把两栏的布局挂起计数留在 >0，子控件会停在旧尺寸。
                Panel1.ResumeLayout(false);
                Panel2.ResumeLayout(false);
                if (_retryQueued || _retryCount >= 5 || IsDisposed || !IsHandleCreated) return;
                _retryCount++;
                _retryQueued = true;
                try
                {
                    BeginInvoke(new Action(RetryLayout));
                }
                catch (InvalidOperationException)
                {
                    _retryQueued = false;
                }
            }
        }

        private void RetryLayout()
        {
            _retryQueued = false;
            if (IsDisposed) return;
            for (int i = 0; i < 3; i++)
            {
                Panel1.ResumeLayout(true);
                Panel2.ResumeLayout(true);
            }
            PerformLayout();
        }
    }
}
