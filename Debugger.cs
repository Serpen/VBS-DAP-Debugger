using Microsoft.VisualStudio.Debugger.Interop;
using static Helpers;

class Debugger : IApplicationDebugger, IDebugSessionProvider
{
    public int QueryAlive()
    {
        System.Diagnostics.Debug.WriteLine("myDebugger.QueryAlive");
        throw new NotImplementedException();
    }

    public int CreateInstanceAtDebugger(ref Guid rclsid, object pUnkOuter, uint dwClsContext, ref Guid riid, out object ppvObject)
    {
        System.Diagnostics.Debug.WriteLine("myDebugger.CreateInstanceAtDebugger");
        throw new NotImplementedException();
    }

    public int onDebugOutput(string pstr)
    {
        Console.WriteLine("onDebugOutput {0}", pstr);
        return 0;
    }

    public int onClose()
    {
        System.Diagnostics.Debug.WriteLine("myDebugger.onClose");
        throw new NotImplementedException();
    }

    public int onDebuggerEvent(ref Guid riid, object punk)
    {
        System.Diagnostics.Debug.WriteLine("myDebugger.onDebuggerEvent");
        throw new NotImplementedException();
    }

    public int StartDebugSession(IRemoteDebugApplication pda)
    {
        System.Diagnostics.Debug.WriteLine("myDebugger.StartDebugSession");
        return SUCCESS(pda.ConnectDebugger(this));
    }


    public int onHandleBreakPoint(IRemoteDebugApplicationThread prpt, BREAKREASON br, IActiveScriptErrorDebug pError)
    {
        Program.vbsbase.DebugThread = prpt;

        SUCCESS(prpt.GetDescription(out var des, out var state));

        SUCCESS(prpt.EnumStackFrames(out var edsf_native));

#if ARCH64
        var enumDebugStackFrames = edsf_native as IEnumDebugStackFrames64 ?? throw new Exception("no IEnumDebugStackFrames");
#else
        var enumDebugStackFrames = edsf_native;
#endif

        SUCCESS(enumDebugStackFrames.Reset());

#if ARCH64
        var dstd = new DebugStackFrameDescriptor64[1];
#else
        var dstd = new DebugStackFrameDescriptor[1];
#endif

#if ARCH64
        SUCCESS(enumDebugStackFrames.Next64((uint)dstd.Length, dstd, out _));
#else
        SUCCESS(enumDebugStackFrames.Next((uint)dstd.Length, dstd, out _));
#endif

        SUCCESS(dstd[0].pdsf.GetCodeContext(out var debugCodeContext));

        SUCCESS(debugCodeContext.GetDocumentContext(out var debugDocumentContext));

        SUCCESS(debugDocumentContext.GetDocument(out var debugDocument));

        var debugDocumentText = debugDocument as IDebugDocumentText ?? throw new Exception("no IDebugDocumentText");

        SUCCESS(debugDocumentText.GetPositionOfContext(debugDocumentContext, out var charpos, out var charnum));

        SUCCESS(debugDocumentText.GetLineOfPosition(charpos, out var line, out var charoffset));

        Console.WriteLine($"{br} {des} {state} ");
        if (br == BREAKREASON.BREAKREASON_ERROR) {
            var exp = new stdole.EXCEPINFO[1];
            pError.GetExceptionInfo(exp);
            System.Console.WriteLine($":{line+1},{charoffset+1} {exp[0].bstrSource}: {exp[0].bstrDescription} &H{exp[0].scode.ToString("X")}");
        }


        return 0;
    }   
}