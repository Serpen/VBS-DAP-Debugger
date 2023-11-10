using Microsoft.VisualStudio.Debugger.Interop;
using System.Runtime.InteropServices;
using static Helpers;
using DAP = Microsoft.VisualStudio.Shared.VSCodeDebugProtocol;

public delegate void CloseHandler();

class Debugger : IApplicationDebugger, IDebugSessionProvider
{
    readonly VbsDebugAdapter Dap;

    StackFrame sf1;
    BREAKREASON lastbr;

    internal stdole.EXCEPINFO[] exp;

    public Debugger(VbsDebugAdapter dap)
    {
        this.Dap = dap;
    }

    public int QueryAlive()
    {
        DebugWriteMethodeName();
        return S_OK;
    }

    public int CreateInstanceAtDebugger(ref Guid rclsid, object pUnkOuter, uint dwClsContext, ref Guid riid, out object ppvObject)
    {
        DebugWriteMethodeName();
        throw new NotImplementedException();
    }

    public int onDebugOutput(string pstr)
    {
        DebugWriteMethodeName(pstr);
        return S_OK;
    }

    public event CloseHandler Close;

    public int onClose()
    {
        DebugWriteMethodeName();
        Close?.Invoke();
        return S_OK;
    }

    public int onDebuggerEvent(ref Guid riid, object punk)
    {
        DebugWriteMethodeName();
        return S_OK;
    }

    public int StartDebugSession(IRemoteDebugApplication pda)
    {
        DebugWriteMethodeName();
        return SUCCESS(pda.ConnectDebugger(this));
    }


    public int onHandleBreakPoint(IRemoteDebugApplicationThread prpt, BREAKREASON br, IActiveScriptErrorDebug pError)
    {
        DebugWriteMethodeName();
        lastbr = br;

        Dap.DebugThread = prpt;

        SUCCESS(prpt.GetDescription(out var des, out var state));

        sf1 = StackFrame.GetFrames(prpt, true).First();

        SUCCESS(sf1.dsf.GetCodeContext(out var debugCodeContext));

        SUCCESS(debugCodeContext.GetDocumentContext(out var debugDocumentContext));

        SUCCESS(debugDocumentContext.GetDocument(out var debugDocument));

        var debugDocumentText = debugDocument as IDebugDocumentText ?? throw new Exception("no IDebugDocumentText");

        SUCCESS(debugDocumentText.GetPositionOfContext(debugDocumentContext, out var charpos, out var charnum));

        SUCCESS(debugDocumentText.GetLineOfPosition(charpos, out var line, out var charoffset));

        // SUCCESS(debugDocumentText.GetSize(out var totalLines, out var totalChars));
        // IntPtr textPtr = Marshal.AllocHGlobal((int)totalChars * 2);
        // uint readChars = 0;
        // SUCCESS(debugDocumentText.GetText(0, textPtr, null, ref readChars, totalChars), "GetText");
        // string text = Marshal.PtrToStringAuto(textPtr, ((int)readChars));
        // Marshal.FreeHGlobal(textPtr);

        if (!Program.isDAP) System.Console.WriteLine($":{line + 1},{charoffset + 1} {br} {des} {state} ");
        if (br == BREAKREASON.BREAKREASON_ERROR)
        {
            exp = new stdole.EXCEPINFO[1];
            SUCCESS(pError.GetExceptionInfo(exp));
            System.Diagnostics.Debug.WriteLine($"{exp[0].bstrSource}: {exp[0].bstrDescription} &H{exp[0].scode.ToString("X")}");
            if (Program.isDAP) this.Dap.Protocol.SendEvent(new DAP.Messages.StoppedEvent()
            {
                Reason = DAP.Messages.StoppedEvent.ReasonValue.Exception,
                ThreadId = 1,
                Text = $"{exp[0].bstrSource}: {exp[0].bstrDescription} &H{exp[0].scode.ToString("X")}"
            });
            else
                System.Console.WriteLine($"{exp[0].bstrSource}: {exp[0].bstrDescription} &H{exp[0].scode.ToString("X")}");
        }
        else
        {
            if (Program.isDAP) this.Dap.Protocol.SendEvent(new DAP.Messages.StoppedEvent()
            {
                Reason = DAP.Messages.StoppedEvent.ReasonValue.InstructionBreakpoint,
                ThreadId = 1
            });
        }


        return S_OK;
    }
}