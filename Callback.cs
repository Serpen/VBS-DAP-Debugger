using Microsoft.VisualStudio.Debugger.Interop;
using static Helpers;

public class Callback : IDebugPortNotify2
{
    volatile int m_refCount;
    readonly int m_hCodeRanEvent;

    private Callback(int hCodeRanEvent)
    {
        m_refCount = 1;
        m_hCodeRanEvent = hCodeRanEvent;
    }
    ~Callback()
    {
        CloseHandle(m_hCodeRanEvent);
    }

    internal Guid irdapp = new Guid("51973C30-CB0C-11D0-B5C9-00A0244A0E7A");
    public int AddProgramNode(IDebugProgramNode2 pProgramNode)
    {
        System.Diagnostics.Debug.WriteLine("AddProgramNode");
        DebugWriteMethodeName();

        var x = pProgramNode as IDebugProviderProgramNode2;
        x.UnmarshalDebuggeeInterface(ref irdapp, out var intPtr);

        Program.vbsbase.CauseBreak();

        return 0;
    }

    public int RemoveProgramNode(IDebugProgramNode2 pProgramNode)
    {
        DebugWriteMethodeName();
        throw new NotImplementedException();
    }

    [System.Runtime.InteropServices.DllImport("ole32.dll")]
    static extern int CoWaitForMultipleHandles(uint dwFlags, uint dwTimeout, int cHandles, int[] pHandles, out int lpdwindex);
    [System.Runtime.InteropServices.DllImport("kernel32.dll")]
    static extern IntPtr CreateEvent(IntPtr lpEventAttributes, bool bManualReset, bool bInitialState, string lpName);
    [System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError = true)]
    [System.Runtime.ConstrainedExecution.ReliabilityContract(System.Runtime.ConstrainedExecution.Consistency.WillNotCorruptState, System.Runtime.ConstrainedExecution.Cer.Success)]
    [System.Security.SuppressUnmanagedCodeSecurity()]
    [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
    static extern bool CloseHandle(IntPtr hObject);

    internal void Wait(int hProcess)
    {
        int[] handles = { hProcess, m_hCodeRanEvent };
        int index = 0;
        CoWaitForMultipleHandles(0, uint.MaxValue, handles.Length, handles, out index);

        if (index == 0)
            System.Console.WriteLine("WARNING: IE exited without running script code.\n");

    }

    internal static int Create(out Callback ppNewObject)
    {
        // DebugWriteMethodeName();
        var hEvent = CreateEvent(IntPtr.Zero, true, false, null);
        if (hEvent == IntPtr.Zero)
        {
            throw new System.ComponentModel.Win32Exception(System.Runtime.InteropServices.Marshal.GetLastWin32Error());
        }

        ppNewObject = new Callback((int)hEvent);
        return 0;
    }
}