using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace wxd.files.winforms;

/// <summary>传输中心：查看排队中的复制/移动任务，支持暂停、继续、取消。</summary>
internal sealed class TransferCenterForm : Form
{
    private static TransferCenterForm? _instance;

    private readonly ListView _list = new();
    private readonly Dictionary<string, ListViewItem> _rows = new(StringComparer.Ordinal);
    private readonly System.Windows.Forms.Timer _timer = new() { Interval = 400 };
    private readonly ThemeButton _btnPause = new();
    private readonly ThemeButton _btnCancel = new();
    private readonly ThemeButton _btnClear = new();
    private readonly ThemeButton _btnClose = new();

    private TransferCenterForm()
    {
        Text = "传输中心";
        ClientSize = new Size(920, 440);
        MinimumSize = new Size(720, 320);
        StartPosition = FormStartPosition.CenterParent;
        ShowIcon = false;
        Font = Theme.Body;

        BuildUi();
        ApplyTheme();
        Theme.Changed += ApplyTheme;
        FileOperationQueue.Shared.JobChanged += OnJobChanged;
        _timer.Tick += (_, _) => RefreshJobs();
        _timer.Start();

        RefreshJobs();
    }

    /// <summary>打开或激活唯一的传输中心窗口。</summary>
    public static void ShowCenter(IWin32Window? owner)
    {
        if (_instance is { IsDisposed: false })
        {
            if (_instance.WindowState == FormWindowState.Minimized)
                _instance.WindowState = FormWindowState.Normal;
            _instance.Activate();
            return;
        }

        _instance = new TransferCenterForm();
        if (owner != null) _instance.Show(owner);
        else _instance.Show();
    }

    private void BuildUi()
    {
        _list.Dock = DockStyle.Fill;
        _list.View = View.Details;
        _list.FullRowSelect = true;
        _list.HideSelection = false;
        _list.MultiSelect = false;
        _list.GridLines = false;
        _list.Columns.Add("状态", 70);
        _list.Columns.Add("类型", 55);
        _list.Columns.Add("来源", 300);
        _list.Columns.Add("目标", 170);
        _list.Columns.Add("进度", 150);
        _list.Columns.Add("信息", 220);
        _list.SelectedIndexChanged += (_, _) => UpdateButtons();

        var bar = new Panel { Dock = DockStyle.Bottom, Height = ThemeMetrics.BarHeight + 8 };
        _btnPause.Text = "暂停";
        _btnCancel.Text = "取消";
        _btnClear.Text = "清除已完成";
        _btnClose.Text = "关闭";
        _btnPause.Click += (_, _) => TogglePause();
        _btnCancel.Click += (_, _) => CancelSelected();
        _btnClear.Click += (_, _) => FileOperationQueue.Shared.ClearFinished();
        _btnClose.Click += (_, _) => Close();
        bar.Controls.AddRange(new Control[] { _btnPause, _btnCancel, _btnClear, _btnClose });
        DialogButtonBar.LayoutRight(bar, _btnPause, _btnCancel, _btnClear, _btnClose);
        bar.Resize += (_, _) =>
            DialogButtonBar.LayoutRight(bar, _btnPause, _btnCancel, _btnClear, _btnClose);

        Controls.Add(_list);
        Controls.Add(bar);
        CancelButton = _btnClose;
    }

    private void TogglePause()
    {
        var job = SelectedJob();
        if (job == null) return;
        if (job.State == FileOperationState.Paused) FileOperationQueue.Shared.Resume(job.Id);
        else FileOperationQueue.Shared.Pause(job.Id);
        RefreshJobs();
    }

    private void CancelSelected()
    {
        var job = SelectedJob();
        if (job == null) return;
        FileOperationQueue.Shared.Cancel(job.Id);
        RefreshJobs();
    }

    private FileOperationJob? SelectedJob()
    {
        if (_list.SelectedItems.Count != 1) return null;
        return _list.SelectedItems[0].Tag is string id ? FileOperationQueue.Shared.Find(id) : null;
    }

    private void OnJobChanged(object? sender, FileOperationJob job)
    {
        if (IsDisposed) return;
        try
        {
            if (IsHandleCreated) BeginInvoke(new Action(RefreshJobs));
        }
        catch
        {
            // 窗口正在关闭
        }
    }

    private void RefreshJobs()
    {
        if (IsDisposed) return;

        var jobs = FileOperationQueue.Shared.Jobs;
        var alive = new HashSet<string>(StringComparer.Ordinal);
        foreach (var job in jobs)
        {
            alive.Add(job.Id);
            if (!_rows.TryGetValue(job.Id, out var item))
            {
                item = new ListViewItem(new[] { job.StateText, job.KindText, "", "", "", "" });
                item.Tag = job.Id;
                _rows[job.Id] = item;
                _list.Items.Add(item);
            }
            UpdateItem(item, job);
        }

        foreach (var id in _rows.Keys.ToList())
        {
            if (alive.Contains(id)) continue;
            _list.Items.Remove(_rows[id]);
            _rows.Remove(id);
        }

        UpdateButtons();
    }

    private static void UpdateItem(ListViewItem item, FileOperationJob job)
    {
        item.SubItems[0].Text = job.StateText;
        item.SubItems[1].Text = job.KindText;
        item.SubItems[2].Text = DescribeSources(job);
        item.SubItems[3].Text = job.Destination ?? "—";
        item.SubItems[4].Text = BuildProgress(job);
        item.SubItems[5].Text = job.Error
            ?? (job.FailedItems > 0
                ? "失败 " + job.FailedItems + " 项"
                : job.CurrentItem != null ? Path.GetFileName(job.CurrentItem) : "");
        item.ForeColor = job.State switch
        {
            FileOperationState.Failed => Theme.Signal(Color.FromArgb(200, 60, 60)),
            FileOperationState.Cancelled => Theme.P.Muted,
            FileOperationState.Completed => Theme.P.Muted,
            _ => Theme.P.Text
        };
    }

    private static string DescribeSources(FileOperationJob job)
    {
        if (job.Sources.Count == 0) return "";
        var first = job.CurrentItem ?? job.Sources[0];
        return job.Sources.Count == 1 ? first : first + " 等 " + job.Sources.Count + " 项";
    }

    private static string BuildProgress(FileOperationJob job)
    {
        var percent = job.Percent.ToString("0");
        var bytes = job.TotalBytes > 0
            ? FormatSize(job.DoneBytes) + " / " + FormatSize(job.TotalBytes)
            : job.DoneItems + " / " + job.TotalItems + " 项";
        return percent + "%  (" + bytes + ")";
    }

    private static string FormatSize(long bytes)
    {
        if (bytes < 1024) return bytes + " B";
        string[] units = { "KB", "MB", "GB", "TB" };
        double value = bytes;
        int unit = -1;
        while (value >= 1024 && unit < units.Length - 1)
        {
            value /= 1024;
            unit++;
        }
        return value.ToString("0.##") + " " + units[unit];
    }

    private void UpdateButtons()
    {
        var job = SelectedJob();
        bool active = job is { State: FileOperationState.Running or FileOperationState.Paused };
        _btnPause.Enabled = active;
        _btnPause.Text = job?.State == FileOperationState.Paused ? "继续" : "暂停";
        _btnCancel.Enabled = job is { State: FileOperationState.Pending or FileOperationState.Running or FileOperationState.Paused };
        _btnClear.Enabled = FileOperationQueue.Shared.Jobs.Any(j =>
            j.State is FileOperationState.Completed or FileOperationState.Failed or FileOperationState.Cancelled);
    }

    private void ApplyTheme()
    {
        var p = Theme.P;
        BackColor = p.Window;
        ForeColor = p.Text;
        _list.BackColor = p.Surface;
        _list.ForeColor = p.Text;
        foreach (Control control in Controls)
        {
            if (control is Panel panel) panel.BackColor = p.Chrome;
        }
        _btnPause.BackColor = p.Chrome;
        _btnCancel.BackColor = p.Chrome;
        _btnClear.BackColor = p.Chrome;
        _btnClose.BackColor = p.Chrome;
        Invalidate(true);
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        _timer.Stop();
        _timer.Dispose();
        FileOperationQueue.Shared.JobChanged -= OnJobChanged;
        Theme.Changed -= ApplyTheme;
        _instance = null;
        base.OnFormClosed(e);
    }
}
