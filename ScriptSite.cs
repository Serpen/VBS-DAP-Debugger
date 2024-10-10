using static Helpers;

class ScriptSite : ActiveDbg.IActiveScriptSite, VSDebug.IActiveScriptSiteDebug64, VSDebug.IActiveScriptSiteDebug32, VSDebug.IActiveScriptSiteWindow
{
    readonly VSDebug.IDebugApplication64 m_debugApplication64;
    readonly DAP.DebugAdapterBase dap;
    readonly VSDebug.IDebugApplication32 m_debugApplication32;
    internal Dictionary<string, object> NamedItems { get; } = new Dictionary<string, object>();
    const uint TYPE_E_ELEMENTNOTFOUND = 0x8002802B;

    internal ScriptSite(VSDebug.IDebugApplication64 debugApplication, DAP.DebugAdapterBase dap)
    {
        m_debugApplication64 = debugApplication;
        this.dap = dap;
    }

    internal ScriptSite(VSDebug.IDebugApplication32 debugApplication, DAP.DebugAdapterBase dap)
    {
        m_debugApplication32 = debugApplication;
        this.dap = dap;
    }

    public int GetDocVersionString(out string version)
    {
        version = new Version().ToString();
        System.Diagnostics.Debug.WriteLine($"{nameof(ScriptSite)}.{nameof(GetDocVersionString)} {version}");
        return S_OK;
    }

    public uint GetItemInfo(string pstrName, uint dwReturnMask, out object ppiunkItem, IntPtr ppti)
    {
        System.Diagnostics.Debug.WriteLine($"{nameof(ScriptSite)}.{nameof(GetItemInfo)} {pstrName} {dwReturnMask}");
        if (NamedItems.TryGetValue(pstrName, out var item))
        {
            ppiunkItem = item;
            return S_OK;
        }
        else
        {
            ppiunkItem = null;
            return TYPE_E_ELEMENTNOTFOUND;
        }
    }

    public int GetLCID(out uint lcid)
    {
        lcid = (uint)System.Globalization.CultureInfo.CurrentUICulture.LCID;
        System.Diagnostics.Debug.WriteLine($"{nameof(ScriptSite)}.{nameof(GetLCID)} {lcid}");
        return S_OK;
    }

    public int OnEnterScript()
    {
        System.Diagnostics.Debug.WriteLine($"{nameof(ScriptSite)}.{nameof(OnEnterScript)}");
        return S_OK;
    }

    public int OnLeaveScript()
    {
        if (Program.isDAP) this.dap.Protocol.SendEvent(new DAP.Messages.ExitedEvent());
        System.Diagnostics.Debug.WriteLine($"{nameof(ScriptSite)}.{nameof(OnLeaveScript)}");
        return S_OK;
    }

    public int OnScriptError(VSDebug.IActiveScriptError scriptError)
    {
        stdole.EXCEPINFO[] expinf = new stdole.EXCEPINFO[1];
        scriptError.GetExceptionInfo(expinf);
        System.Diagnostics.Debug.WriteLine($"{nameof(ScriptSite)}.{nameof(OnScriptError)} {scriptError} {expinf[0].bstrDescription} {expinf[0].bstrSource} {expinf[0].scode.ToString("X")}");
        return 0;
    }

    public int OnScriptTerminate(ref object result, stdole.EXCEPINFO[] exceptionInfo)
    {
        System.Diagnostics.Debug.WriteLine($"{nameof(ScriptSite)}.{nameof(OnScriptTerminate)} {result}");
        return S_OK;
    }

    public int OnStateChange(VSDebug.SCRIPTSTATE scriptState)
    {
        System.Diagnostics.Debug.WriteLine($"{nameof(ScriptSite)}.{nameof(OnStateChange)} {scriptState}");
        return S_OK;
    }
    int VSDebug.IActiveScriptSiteDebug64.GetDocumentContextFromPosition(ulong dwSourceContext, uint uCharacterOffset, uint uNumChars, out VSDebug.IDebugDocumentContext ppsc)
    {
        System.Diagnostics.Debug.WriteLine($"{nameof(ScriptSite)}64.{nameof(VSDebug.IActiveScriptSiteDebug64.GetDocumentContextFromPosition)} {dwSourceContext}, {uCharacterOffset} {uNumChars}");
        throw new NotImplementedException();
    }
    int VSDebug.IActiveScriptSiteDebug32.GetDocumentContextFromPosition(uint dwSourceContext, uint uCharacterOffset, uint uNumChars, out VSDebug.IDebugDocumentContext ppsc)
    {
        System.Diagnostics.Debug.WriteLine($"{nameof(ScriptSite)}32.{nameof(VSDebug.IActiveScriptSiteDebug32.GetDocumentContextFromPosition)} {dwSourceContext}, {uCharacterOffset} {uNumChars}");
        throw new NotImplementedException();
    }

    int VSDebug.IActiveScriptSiteDebug32.GetApplication(out VSDebug.IDebugApplication32 ppda)
    {
        DebugWriteMethodeName();
        ppda = m_debugApplication32;
        return S_OK;
    }
    int VSDebug.IActiveScriptSiteDebug64.GetApplication(out VSDebug.IDebugApplication64 ppda)
    {
        DebugWriteMethodeName();
        ppda = m_debugApplication64;
        return S_OK;
    }

    int VSDebug.IActiveScriptSiteDebug64.GetRootApplicationNode(out VSDebug.IDebugApplicationNode ppdanRoot)
    {
        SUCCESS(m_debugApplication64.GetRootNode(out ppdanRoot));
        SUCCESS(ppdanRoot.GetName(VSDebug.DOCUMENTNAMETYPE.DOCUMENTNAMETYPE_TITLE, out var title));
        DebugWriteMethodeName();
        return 0x8004001;
    }
    int VSDebug.IActiveScriptSiteDebug32.GetRootApplicationNode(out VSDebug.IDebugApplicationNode ppdanRoot)
    {
        SUCCESS(m_debugApplication32.GetRootNode(out ppdanRoot));
        SUCCESS(ppdanRoot.GetName(VSDebug.DOCUMENTNAMETYPE.DOCUMENTNAMETYPE_TITLE, out var title));
        DebugWriteMethodeName();
        return 0x8004001;
    }

    public int OnScriptErrorDebug(VSDebug.IActiveScriptErrorDebug pErrorDebug, out int pfEnterDebugger, out int pfCallOnScriptErrorWhenContinuing)
    {
        DebugWriteMethodeName();
        pfEnterDebugger = 0;
        pfCallOnScriptErrorWhenContinuing = 1;
        return S_OK;
    }

    public int GetWindow(out IntPtr phwnd)
    {
        phwnd = IntPtr.Zero;
        System.Diagnostics.Debug.WriteLine($"{nameof(ScriptSite)}.{nameof(GetWindow)} {phwnd}");
        return S_OK;
    }

    public int EnableModeless(int fEnable)
    {
        System.Diagnostics.Debug.WriteLine($"{nameof(ScriptSite)}.{nameof(EnableModeless)} {fEnable}");
        return S_OK;
    }
}
