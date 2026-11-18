using System;
using System.Collections.Generic;
using System.Linq;

namespace wxd.files.winforms;

/// <summary>
/// 把传输任务的原始异常翻译成用户能看懂、能行动的中文提示。
/// 权限类错误只显示英文的 "Access to the path ... is denied."，用户不知道下一步做什么。
/// </summary>
internal static class OperationErrorPresenter
{
    private static readonly string[] AccessDeniedMarkers =
    {
        "Access to the path",
        "is denied",
        "Access is denied",
        "UnauthorizedAccess",
        "拒绝访问",
        "访问被拒绝",
        "没有权限"
    };

    public static bool IsAccessDenied(IEnumerable<string> errors)
    {
        foreach (var error in errors)
        {
            if (string.IsNullOrEmpty(error)) continue;
            foreach (var marker in AccessDeniedMarkers)
            {
                if (error.Contains(marker, StringComparison.OrdinalIgnoreCase)) return true;
            }
        }
        return false;
    }

    public static string Build(FileOperationJob job)
    {
        var lines = new List<string>
        {
            job.KindText + "失败：" + job.FailedItems + " 项未完成。"
        };

        if (!string.IsNullOrEmpty(job.Destination))
            lines.Add("目标位置：" + job.Destination);

        var first = job.Errors.FirstOrDefault() ?? job.Error ?? "未知错误";

        if (IsAccessDenied(job.Errors) || IsAccessDenied(new[] { job.Error ?? "" }))
        {
            lines.Add("");
            lines.Add("系统拒绝了写入：目标目录没有写入权限（常见于磁盘根目录、系统目录、只读目录，或被其他程序占用）。");
            lines.Add("");
            lines.Add("请换到有写入权限的目录后重试；如果确实要写系统目录，请右键程序 →「以管理员身份运行」。");
            lines.Add("");
            lines.Add("详细信息：" + first);
        }
        else
        {
            lines.Add("");
            lines.Add(first);
        }

        return string.Join(GlobalConst.NewLine, lines);
    }
}
