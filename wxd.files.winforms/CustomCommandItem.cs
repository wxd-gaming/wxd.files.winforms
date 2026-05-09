using System.Drawing;

namespace wxd.files.winforms;

/// <summary>自定义命令的数据模型</summary>
public class CustomCommandItem
{
    /// <summary>命令显示名称（按钮上的文字）</summary>
    public string Name { get; set; } = "";

    /// <summary>要执行的命令（exe路径 或 cmd /c xxx）</summary>
    public string Command { get; set; } = "";

    /// <summary>命令执行的工作目录（留空则使用默认目录）</summary>
    public string WorkDir { get; set; } = "";

    /// <summary>是否显示 cmd 窗口（false=后台静默执行，true=弹出窗口可看到输出）</summary>
    public bool ShowWindow { get; set; } = false;

    /// <summary>按钮颜色（ARGB 十六进制字符串，如 "#7B1FA2"），便于分组</summary>
    public string ButtonColor { get; set; } = "#7B1FA2";
}
