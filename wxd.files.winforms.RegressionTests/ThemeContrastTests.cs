using System;
using System.Drawing;
using System.Reflection;
using System.Threading;

namespace wxd.files.winforms.RegressionTests;

/// <summary>
/// 选中行可读性回归。
/// 背景：日志查询等自绘选中行曾拿 AccentText 当选中文字色，而 AccentText 是给饱和 Accent 底配的白字，
/// 浅色皮肤的选中底是淡彩，白字对比度只有 1.1~1.5:1，选中整行看不清。
/// 守两件事：所有皮肤的 SelectionText 对 Selection 达到 AA 4.5:1；日志列表确实取这个令牌且切肤后跟随。
/// </summary>
internal static partial class Program
{
    private static void RunThemeContrastTests()
    {
        Console.WriteLine("--- 主题选中行可读性 ---");
        Run("主题 所有皮肤选中文字对比度≥4.5", Test_SelectionTextContrastAllPalettes);
        Run("主题 日志列表选中色取 SelectionText", Test_LineListViewSelectionTextFollowsTheme);
    }

    /// <summary>WCAG 2.x 相对亮度对比度；返回 1~21。</summary>
    private static double ContrastRatio(Color fg, Color bg)
    {
        static double Luminance(Color c)
        {
            static double Channel(int v)
            {
                double s = v / 255.0;
                return s <= 0.03928 ? s / 12.92 : Math.Pow((s + 0.055) / 1.055, 2.4);
            }
            return 0.2126 * Channel(c.R) + 0.7152 * Channel(c.G) + 0.0722 * Channel(c.B);
        }

        double a = Luminance(fg);
        double b = Luminance(bg);
        return (Math.Max(a, b) + 0.05) / (Math.Min(a, b) + 0.05);
    }

    private static void Test_SelectionTextContrastAllPalettes()
    {
        var original = (Theme.Skin, Theme.Dark);
        try
        {
            foreach (UiSkin skin in Enum.GetValues<UiSkin>())
            {
                foreach (bool dark in new[] { false, true })
                {
                    // 与当前皮肤相同时 Set 直接返回，Theme.P 已是该套配色，照测即可
                    Theme.Set(skin, dark);
                    AssertSelectionContrast(skin, dark);
                }
            }
        }
        finally
        {
            Theme.Set(original.Skin, original.Dark);
        }
    }

    private static void AssertSelectionContrast(UiSkin skin, bool dark)
    {
        var p = Theme.P;
        double ratio = ContrastRatio(p.SelectionText, p.Selection);
        Assert(ratio >= 4.5,
            $"{skin}/{(dark ? "深" : "浅")} 选中文字 #{p.SelectionText.R:X2}{p.SelectionText.G:X2}{p.SelectionText.B:X2}" +
            $" 压选中底 #{p.Selection.R:X2}{p.Selection.G:X2}{p.Selection.B:X2} 对比度 {ratio:0.00}:1，低于 4.5:1");
    }

    private static void Test_LineListViewSelectionTextFollowsTheme()
    {
        Exception? captured = null;
        var thread = new Thread(() =>
        {
            var original = (Theme.Skin, Theme.Dark);
            try
            {
                var field = typeof(LineListView).GetField("_selectionText", BindingFlags.NonPublic | BindingFlags.Instance);
                Assert(field != null, "LineListView._selectionText 字段应存在");

                // 截图现场是 Mint 浅色：白字压淡薄荷底，整行读不出来
                Theme.Set(UiSkin.Mint, false);
                using (var list = new LineListView())
                {
                    var mintText = (Color)field!.GetValue(list)!;
                    AssertEq(Theme.P.SelectionText.ToArgb(), mintText.ToArgb(), "创建时应取 SelectionText");
                    Assert(ContrastRatio(mintText, Theme.P.Selection) >= 4.5, "Mint 浅色选中行必须可读");

                    Theme.Set(UiSkin.Sky, true);
                    list.ApplyTheme();
                    var darkText = (Color)field.GetValue(list)!;
                    AssertEq(Theme.P.SelectionText.ToArgb(), darkText.ToArgb(), "切肤后应重新取 SelectionText");
                    Assert(ContrastRatio(darkText, Theme.P.Selection) >= 4.5, "深色选中行必须可读");
                }
            }
            catch (Exception ex)
            {
                captured = ex;
            }
            finally
            {
                Theme.Set(original.Skin, original.Dark);
            }
        })
        {
            IsBackground = true
        };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();

        Assert(thread.Join(15000), "LineListView 主题测试线程超时");
        if (captured != null) throw new Exception("LineListView 选中色测试失败: " + captured);
    }
}