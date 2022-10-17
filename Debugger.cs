using Microsoft.VisualStudio.Debugger.Interop;
using System.Runtime.InteropServices;
using static Helpers;

public delegate void CloseHandler();

class Debugger : IApplicationDebugger, IDebugSessionProvider
{
    public int QueryAlive()
    {
        System.Diagnostics.Debug.WriteLine($"{nameof(Debugger)}.{nameof(QueryAlive)}");
        return S_OK;
    }

    public int CreateInstanceAtDebugger(ref Guid rclsid, object pUnkOuter, uint dwClsContext, ref Guid riid, out object ppvObject)
    {
        System.Diagnostics.Debug.WriteLine($"{nameof(Debugger)}.{nameof(CreateInstanceAtDebugger)}");
        throw new NotImplementedException();
    }

    public int onDebugOutput(string pstr)
    {
        System.Diagnostics.Debug.WriteLine($"{nameof(Debugger)}.{nameof(onDebugOutput)}: {pstr}");
        return S_OK;
    }

    public event CloseHandler? Close;

    public int onClose()
    {
        System.Diagnostics.Debug.WriteLine($"{nameof(Debugger)}.{nameof(onClose)}");
        Close?.Invoke();
        return S_OK;
    }

    public int onDebuggerEvent(ref Guid riid, object punk)
    {
        System.Diagnostics.Debug.WriteLine($"{nameof(Debugger)}.{nameof(onDebuggerEvent)}");
        return S_OK;
    }

    public int StartDebugSession(IRemoteDebugApplication pda)
    {
        System.Diagnostics.Debug.WriteLine($"{nameof(Debugger)}.{nameof(StartDebugSession)}");
        return SUCCESS(pda.ConnectDebugger(this));
    }

    public int onHandleBreakPoint(IRemoteDebugApplicationThread prpt, BREAKREASON br, IActiveScriptErrorDebug pError)
    {
        System.Diagnostics.Debug.WriteLine($"{nameof(Debugger)}.{nameof(onHandleBreakPoint)}");

        Program.vbsbase.DebugThread = prpt;

        SUCCESS(prpt.GetDescription(out var des, out var state));

        var sf1 = StackFrame.GetFrames(prpt, true).First();

        SUCCESS(sf1.dsf.GetCodeContext(out var debugCodeContext));

        SUCCESS(debugCodeContext.GetDocumentContext(out var debugDocumentContext));

        SUCCESS(debugDocumentContext.GetDocument(out var debugDocument));

        var debugDocumentText = debugDocument as IDebugDocumentText ?? throw new Exception("no IDebugDocumentText");

        SUCCESS(debugDocumentText.GetPositionOfContext(debugDocumentContext, out var charpos, out var charnum));

        SUCCESS(debugDocumentText.GetLineOfPosition(charpos, out var line, out var charoffset));

        debugDocumentText.GetSize(out var totalLines, out var totalChars);

        string text = "";
        try
        {
            IntPtr textPtr = Marshal.AllocHGlobal((int)totalChars);
            debugDocumentText.GetText(0, textPtr, new ushort[0], totalChars, totalChars); // TODO: ending AND encondig
            text = Marshal.PtrToStringAuto(textPtr);
        }
        catch (Exception) { }

        Console.WriteLine($":{line + 1},{charoffset + 1} {br} {des} {state} ");
        if (br == BREAKREASON.BREAKREASON_ERROR)
        {
            var exp = new stdole.EXCEPINFO[1];
            pError.GetExceptionInfo(exp);
            System.Console.WriteLine($"{exp[0].bstrSource}: {exp[0].bstrDescription} &H{exp[0].scode.ToString("X")}");
        }

        return S_OK;
    }
}