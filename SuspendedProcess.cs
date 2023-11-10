/*
* just a simple conversion of http://www.codeproject.com/Articles/230005/Launch-a-process-suspended into a command line app
*/

using System.Diagnostics;
using System.Runtime.InteropServices;

[Flags]
public enum ProcessCreationFlags : uint
{
    ZERO_FLAG = 0x00000000,
    CREATE_BREAKAWAY_FROM_JOB = 0x01000000,
    CREATE_DEFAULT_ERROR_MODE = 0x04000000,
    CREATE_NEW_CONSOLE = 0x00000010,
    CREATE_NEW_PROCESS_GROUP = 0x00000200,
    CREATE_NO_WINDOW = 0x08000000,
    CREATE_PROTECTED_PROCESS = 0x00040000,
    CREATE_PRESERVE_CODE_AUTHZ_LEVEL = 0x02000000,
    CREATE_SEPARATE_WOW_VDM = 0x00001000,
    CREATE_SHARED_WOW_VDM = 0x00001000,
    CREATE_SUSPENDED = 0x00000004,
    CREATE_UNICODE_ENVIRONMENT = 0x00000400,
    DEBUG_ONLY_THIS_PROCESS = 0x00000002,
    DEBUG_PROCESS = 0x00000001,
    DETACHED_PROCESS = 0x00000008,
    EXTENDED_STARTUPINFO_PRESENT = 0x00080000,
    INHERIT_PARENT_AFFINITY = 0x00010000
}

public struct PROCESS_INFORMATION
{
    public IntPtr hProcess;
    public IntPtr hThread;
    public uint dwProcessId;
    public uint dwThreadId;
}

public struct STARTUPINFO
{
    public uint cb;
    public string lpReserved;
    public string lpDesktop;
    public string lpTitle;
    public uint dwX;
    public uint dwY;
    public uint dwXSize;
    public uint dwYSize;
    public uint dwXCountChars;
    public uint dwYCountChars;
    public uint dwFillAttribute;
    public uint dwFlags;
    public short wShowWindow;
    public short cbReserved2;
    public IntPtr lpReserved2;
    public IntPtr hStdInput;
    public IntPtr hStdOutput;
    public IntPtr hStdError;
}




public class SuspendedProcess
{

    [DllImport("kernel32.dll", SetLastError=true, CharSet=CharSet.Auto)]
    static extern bool CreateProcess(string lpApplicationName, string lpCommandLine, IntPtr lpProcessAttributes, IntPtr lpThreadAttributes,
                             bool bInheritHandles, ProcessCreationFlags dwCreationFlags, IntPtr lpEnvironment,
                            string lpCurrentDirectory, ref STARTUPINFO lpStartupInfo, out PROCESS_INFORMATION lpProcessInformation);

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern uint ResumeThread(IntPtr hThread);

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern uint SuspendThread(IntPtr hThread);


    readonly IntPtr ThreadHandle = IntPtr.Zero;
    internal readonly Process Process;

    private SuspendedProcess(uint pid, IntPtr threadHandle)
    {
        Process = Process.GetProcessById(((int)pid));
        ThreadHandle = threadHandle;
    }

    public static SuspendedProcess LaunchProcessSuspended(System.Diagnostics.ProcessStartInfo psi)
    {
        STARTUPINFO si = new STARTUPINFO();
        PROCESS_INFORMATION pi = new PROCESS_INFORMATION();
        bool success = CreateProcess(psi.FileName, psi.Arguments, IntPtr.Zero, IntPtr.Zero, true, ProcessCreationFlags.CREATE_SUSPENDED | ProcessCreationFlags.DETACHED_PROCESS, IntPtr.Zero, null, ref si, out pi);
        if (!success)
            throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
        return new SuspendedProcess(pi.dwProcessId, pi.hThread);
    }

    public void ResumeProcess()
    {
        ResumeThread(ThreadHandle);
    }
    public void SuspendProcess()
    {
        SuspendThread(ThreadHandle);
    }
}

