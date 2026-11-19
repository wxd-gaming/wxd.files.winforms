using System;
using System.Runtime.InteropServices;

namespace wxd.files.winforms;

/// <summary>
/// 进程完整性级别。TortoiseSVN 等 Shell 扩展在 Low Integrity 进程里拿不到路径，
/// 右键“SVN 更新”会报“没有在命令行指定路径或文件”。
/// </summary>
internal static class ProcessIntegrity
{
    private const uint TokenQuery = 0x0008;
    private const int TokenIntegrityLevel = 25;

    public static bool IsLow { get; } = DetectLow();

    private static bool DetectLow()
    {
        IntPtr token = IntPtr.Zero;
        IntPtr buffer = IntPtr.Zero;
        try
        {
            if (!OpenProcessToken(GetCurrentProcess(), TokenQuery, out token)) return false;

            buffer = Marshal.AllocHGlobal(64);
            if (!GetTokenInformation(token, TokenIntegrityLevel, buffer, 64, out _)) return false;

            IntPtr sid = Marshal.ReadIntPtr(buffer);
            if (!ConvertSidToStringSid(sid, out IntPtr sidString)) return false;

            var text = Marshal.PtrToStringUni(sidString) ?? "";
            if (sidString != IntPtr.Zero) LocalFree(sidString);
            return text.EndsWith("-4096", StringComparison.Ordinal);
        }
        catch
        {
            return false;
        }
        finally
        {
            if (buffer != IntPtr.Zero) Marshal.FreeHGlobal(buffer);
            if (token != IntPtr.Zero) CloseHandle(token);
        }
    }

    [DllImport("kernel32.dll")]
    private static extern IntPtr GetCurrentProcess();

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool CloseHandle(IntPtr handle);

    [DllImport("advapi32.dll", SetLastError = true)]
    private static extern bool OpenProcessToken(IntPtr process, uint access, out IntPtr token);

    [DllImport("advapi32.dll", SetLastError = true)]
    private static extern bool GetTokenInformation(IntPtr token, int infoClass, IntPtr info,
        int infoLength, out int returnLength);

    [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool ConvertSidToStringSid(IntPtr sid, out IntPtr stringSid);

    [DllImport("kernel32.dll")]
    private static extern IntPtr LocalFree(IntPtr handle);
}
