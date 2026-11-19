using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace wxd.files.winforms;

/// <summary>文件搜索窗口：名称通配/正则、内容行搜索、大小与时间过滤，结果可跳转定位。</summary>
internal sealed class FileSearchForm : Form
{
    private readonly string _startDirectory;
    private readonly Action<string> _onNavigate;

    private readonly TextBox _txtDir = new();
    private readonly TextBox _txtName = new();
    private readonly TextBox _txtContent = new();
    private readonly CheckBox _chkNameRegex = new();
    private readonly CheckBox _chkContentRegex = new();
    private readonly CheckBox _chkSubdirs = new();
    private readonly CheckBox _chkHidden = new();
    private readonly NumericUpDown _numMinSize = new();
    private readonly NumericUpDown _numMaxSize = new();
    private readonly DateTimePicker _dtpFrom = new();
    private readonly DateTimePicker _dtpTo = new();
    private readonly ListView _list = new();
    private readonly Label _lblStatus = new();
    private readonly ThemeButton _btnBrowse = new();
    private readonly ThemeButton _btnSearch = new();
    private readonly ThemeButton _btnCancel = new();
    private readonly ThemeButton _btnClose = new();

    private CancellationTokenSource? _cts;
    private bool _running;

    /// <summary>修改时间过滤是否启用；默认关闭，供 UI 冒烟测试断言。</summary>
    internal bool IsModifiedTimeFilterEnabled => _dtpFrom.Checked || _dtpTo.Checked;

    public FileSearchForm(string startDirectory, Action<string> onNavigate)
    {
        _startDirectory = string.IsNullOrWhiteSpace(startDirectory)
            ? Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
            : startDirectory;
        _onNavigate = onNavigate;

        Text = "搜索文件";
        ClientSize = new Size(960, 600);
        MinimumSize = new Size(760, 420);
        StartPosition = FormStartPosition.CenterParent;
        ShowIcon = false;
        Font = Theme.Body;

        BuildUi();
        ApplyTheme();
        Theme.Changed += ApplyTheme;
        Disposed += (_, _) => Theme.Changed -= ApplyTheme;

        _txtDir.Text = _startDirectory;
        FormClosed += (_, _) =>
        {
            try { _cts?.Cancel(); } catch { }
        };
    }

    private void BuildUi()
    {
        var top = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 104,
            ColumnCount = 3,
            Padding = new Padding(ThemeMetrics.Margin, ThemeMetrics.Gap, ThemeMetrics.Margin, 0)
        };
        top.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 56));
        top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        top.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 96));

        _txtDir.Dock = DockStyle.Fill;
        _txtName.Dock = DockStyle.Fill;
        _txtContent.Dock = DockStyle.Fill;

        _chkNameRegex.Text = "正则";
        _chkNameRegex.AutoSize = true;
        _chkNameRegex.Anchor = AnchorStyles.Left;
        _chkContentRegex.Text = "正则";
        _chkContentRegex.AutoSize = true;
        _chkContentRegex.Anchor = AnchorStyles.Left;

        _btnBrowse.Text = "浏览…";
        _btnBrowse.Size = new Size(80, ThemeMetrics.RowH);
        _btnBrowse.Anchor = AnchorStyles.Left;
        _btnBrowse.Click += (_, _) => BrowseDirectory();

        AddRow(top, 0, "目录", _txtDir, _btnBrowse);
        AddRow(top, 1, "名称", _txtName, _chkNameRegex);
        AddRow(top, 2, "内容", _txtContent, _chkContentRegex);

        var options = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 38,
            WrapContents = false,
            Padding = new Padding(ThemeMetrics.Margin, 2, ThemeMetrics.Margin, 2)
        };

        _chkSubdirs.Text = "含子目录";
        _chkSubdirs.Checked = true;
        _chkSubdirs.AutoSize = true;
        _chkHidden.Text = "含隐藏";
        _chkHidden.AutoSize = true;

        _numMinSize.Minimum = 0;
        _numMinSize.Maximum = 1_000_000_000_000;
        _numMinSize.Width = 84;
        _numMinSize.ThousandsSeparator = true;
        _numMaxSize.Minimum = 0;
        _numMaxSize.Maximum = 1_000_000_000_000;
        _numMaxSize.Width = 84;
        _numMaxSize.ThousandsSeparator = true;

        _dtpFrom.Format = DateTimePickerFormat.Short;
        _dtpFrom.ShowCheckBox = true;
        _dtpFrom.Checked = false;
        _dtpFrom.Width = 112;
        _dtpTo.Format = DateTimePickerFormat.Short;
        _dtpTo.ShowCheckBox = true;
        _dtpTo.Checked = false;
        _dtpTo.Width = 112;

        options.Controls.AddRange(new Control[]
        {
            _chkSubdirs, _chkHidden,
            NewSpacer(10),
            new Label { Text = "大小(字节) ≥", AutoSize = true, Padding = new Padding(0, 6, 0, 0) },
            _numMinSize,
            new Label { Text = "≤", AutoSize = true, Padding = new Padding(0, 6, 0, 0) },
            _numMaxSize,
            NewSpacer(10),
            new Label { Text = "修改时间", AutoSize = true, Padding = new Padding(0, 6, 0, 0) },
            _dtpFrom,
            new Label { Text = "~", AutoSize = true, Padding = new Padding(0, 6, 0, 0) },
            _dtpTo
        });

        var header = new Panel { Dock = DockStyle.Top, Height = 144 };
        header.Controls.Add(options);
        header.Controls.Add(top);

        _list.Dock = DockStyle.Fill;
        _list.View = View.Details;
        _list.FullRowSelect = true;
        _list.HideSelection = false;
        _list.MultiSelect = true;
        _list.Columns.Add("名称", 220);
        _list.Columns.Add("位置", 320);
        _list.Columns.Add("大小", 90, HorizontalAlignment.Right);
        _list.Columns.Add("修改时间", 140);
        _list.Columns.Add("行", 52, HorizontalAlignment.Right);
        _list.Columns.Add("预览", 320);
        _list.DoubleClick += (_, _) => NavigateSelected();

        var bottom = new Panel { Dock = DockStyle.Bottom, Height = ThemeMetrics.BarHeight };
        _lblStatus.Dock = DockStyle.Fill;
        _lblStatus.TextAlign = ContentAlignment.MiddleLeft;
        _lblStatus.Padding = new Padding(ThemeMetrics.Margin, 0, 0, 0);
        _lblStatus.Text = "输入条件后回车开始搜索";

        var buttons = new FlowLayoutPanel
        {
            Dock = DockStyle.Right,
            Width = 330,
            FlowDirection = FlowDirection.RightToLeft,
            Padding = new Padding(0, ThemeMetrics.RowPad, ThemeMetrics.Margin, 0)
        };
        _btnSearch.Text = "搜索";
        _btnSearch.Primary = true;
        _btnSearch.Size = new Size(82, ThemeMetrics.RowH);
        _btnCancel.Text = "取消";
        _btnCancel.Size = new Size(82, ThemeMetrics.RowH);
        _btnCancel.Enabled = false;
        _btnClose.Text = "关闭";
        _btnClose.Size = new Size(82, ThemeMetrics.RowH);

        _btnSearch.Click += (_, _) => _ = RunSearchAsync();
        _btnCancel.Click += (_, _) => CancelSearch();
        _btnClose.Click += (_, _) => Close();

        buttons.Controls.AddRange(new Control[] { _btnClose, _btnCancel, _btnSearch });
        bottom.Controls.Add(_lblStatus);
        bottom.Controls.Add(buttons);

        Controls.Add(_list);
        Controls.Add(header);
        Controls.Add(bottom);

        AcceptButton = _btnSearch;
        CancelButton = _btnClose;

        _txtName.KeyDown += (_, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; _ = RunSearchAsync(); } };
        _txtContent.KeyDown += (_, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; _ = RunSearchAsync(); } };
    }

    private static Control NewSpacer(int width) => new Panel { Width = width, Height = 1 };

    private static void AddRow(TableLayoutPanel grid, int row, string caption, Control field, Control trailing)
    {
        grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));

        var label = new Label
        {
            Text = caption,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
            AutoSize = false
        };

        grid.Controls.Add(label, 0, row);
        grid.Controls.Add(field, 1, row);
        grid.Controls.Add(trailing, 2, row);
    }

    private void BrowseDirectory()
    {
        using var dialog = new FolderBrowserDialog { SelectedPath = _txtDir.Text };
        if (dialog.ShowDialog(this) == DialogResult.OK)
            _txtDir.Text = dialog.SelectedPath;
    }

    private void CancelSearch()
    {
        try { _cts?.Cancel(); } catch { }
    }

    private FileSearchQuery? BuildQuery(out string? error)
    {
        error = null;

        var roots = _txtDir.Text
            .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(p => p.Trim('"'))
            .Where(p => Directory.Exists(p) || File.Exists(p))
            .ToArray();

        if (roots.Length == 0)
        {
            error = "目录不存在，请检查路径。";
            return null;
        }

        var query = new FileSearchQuery
        {
            Roots = roots,
            Pattern = _txtName.Text.Trim(),
            UseRegex = _chkNameRegex.Checked,
            ContentPattern = _txtContent.Text.Trim(),
            ContentRegex = _chkContentRegex.Checked,
            IncludeSubdirectories = _chkSubdirs.Checked,
            IncludeHidden = _chkHidden.Checked,
            MinSize = (long)_numMinSize.Value,
            MaxSize = (long)_numMaxSize.Value,
            ModifiedAfter = _dtpFrom.Checked ? (DateTime?)_dtpFrom.Value : null,
            ModifiedBefore = _dtpTo.Checked ? (DateTime?)_dtpTo.Value : null,
            MaxResults = 5000
        };

        if (query.MinSize > 0 && query.MaxSize > 0 && query.MinSize > query.MaxSize)
        {
            error = "最小大小不能大于最大大小。";
            return null;
        }

        return query;
    }

    private async Task RunSearchAsync()
    {
        if (_running) return;

        var query = BuildQuery(out var error);
        if (query == null)
        {
            if (!string.IsNullOrEmpty(error))
                AlertForm.Show(this, error, "搜索", AlertType.Warning);
            return;
        }

        _cts?.Dispose();
        _cts = new CancellationTokenSource();

        SetRunning(true);
        _list.Items.Clear();
        _lblStatus.Text = "正在搜索…";

        var progress = new Progress<FileSearchProgress>(p =>
        {
            if (!_running) return;
            _lblStatus.Text = "已扫描 " + p.Scanned + " 个文件，命中 " + p.Matched + " 条"
                + (string.IsNullOrEmpty(p.Current) ? "" : "  ·  " + p.Current);
        });

        try
        {
            var result = await FileSearchService.SearchAsync(query, progress, _cts.Token);
            PopulateResults(result);
        }
        catch (OperationCanceledException)
        {
            _lblStatus.Text = "已取消";
        }
        catch (Exception ex)
        {
            _lblStatus.Text = "搜索失败";
            AlertForm.Show(this, "搜索失败: " + ex.Message, "搜索", AlertType.Error, exception: ex);
        }
        finally
        {
            SetRunning(false);
        }
    }

    private void PopulateResults(FileSearchResult result)
    {
        _list.BeginUpdate();
        try
        {
            _list.Items.Clear();
            foreach (var hit in result.Hits)
            {
                var item = new ListViewItem(new[]
                {
                    hit.Name,
                    hit.Directory,
                    hit.IsDirectory ? "" : FormatSize(hit.Size),
                    hit.ModifiedUtc == default ? "" : hit.ModifiedUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm"),
                    hit.LineNumber > 0 ? hit.LineNumber.ToString() : "",
                    hit.Preview ?? ""
                })
                {
                    Tag = hit
                };
                _list.Items.Add(item);
            }
        }
        finally
        {
            _list.EndUpdate();
        }

        _lblStatus.Text = "扫描 " + result.FilesScanned + " 个文件，命中 " + result.Hits.Count + " 条"
            + (result.Truncated ? "（已达上限，请缩小范围）" : "");

        if (result.Errors.Count > 0)
            _lblStatus.Text += "  ·  " + result.Errors.Count + " 个目录读取失败";
    }

    private void NavigateSelected()
    {
        if (_list.SelectedItems.Count == 0) return;
        if (_list.SelectedItems[0].Tag is FileSearchHit hit)
            _onNavigate(hit.Path);
    }

    private void SetRunning(bool running)
    {
        _running = running;
        _btnSearch.Enabled = !running;
        _btnCancel.Enabled = running;
        _txtDir.Enabled = !running;
        _txtName.Enabled = !running;
        _txtContent.Enabled = !running;
        _chkNameRegex.Enabled = !running;
        _chkContentRegex.Enabled = !running;
        _chkSubdirs.Enabled = !running;
        _chkHidden.Enabled = !running;
        _numMinSize.Enabled = !running;
        _numMaxSize.Enabled = !running;
        _dtpFrom.Enabled = !running;
        _dtpTo.Enabled = !running;
        UseWaitCursor = running;
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

    private void ApplyTheme()
    {
        var p = Theme.P;
        BackColor = p.Window;
        ForeColor = p.Text;
        _txtDir.BackColor = p.Surface;
        _txtDir.ForeColor = p.Text;
        _txtName.BackColor = p.Surface;
        _txtName.ForeColor = p.Text;
        _txtContent.BackColor = p.Surface;
        _txtContent.ForeColor = p.Text;
        _list.BackColor = p.Surface;
        _list.ForeColor = p.Text;
        _lblStatus.ForeColor = p.Muted;

        foreach (Control control in AllControls(this))
        {
            if (control is Label label) label.ForeColor = p.Text;
            else if (control is CheckBox check)
            {
                check.ForeColor = p.Text;
                check.BackColor = Color.Transparent;
            }
            else if (control is FlowLayoutPanel flow) flow.BackColor = p.Window;
            else if (control is TableLayoutPanel table) table.BackColor = p.Window;
            else if (control is Panel panel) panel.BackColor = p.Window;
            else if (control is NumericUpDown number)
            {
                number.BackColor = p.Surface;
                number.ForeColor = p.Text;
            }
            else if (control is DateTimePicker picker)
            {
                picker.CalendarMonthBackground = p.Surface;
                picker.CalendarForeColor = p.Text;
            }
            else if (control is ThemeButton button)
            {
                button.BackColor = p.Chrome;
                button.ForeColor = p.Text;
            }
        }

        _lblStatus.ForeColor = p.Muted;
    }

    private static System.Collections.Generic.IEnumerable<Control> AllControls(Control root)
    {
        foreach (Control child in root.Controls)
        {
            yield return child;
            foreach (var nested in AllControls(child)) yield return nested;
        }
    }
}
