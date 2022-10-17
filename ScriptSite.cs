using Microsoft.VisualStudio.Debugger.Interop;
using static Helpers;

class ScriptSite : ActiveDbg.IActiveScriptSiteMy, IActiveScriptSiteDebug64, IActiveScriptSiteDebug32, IActiveScriptSiteWindow
{
    public ScriptSite(IDebugApplication64 debugApplication) => m_debugApplication64 = debugApplication;
    public ScriptSite(IDebugApplication32 debugApplication) => m_debugApplication32 = debugApplication;

    internal readonly IDebugApplication64 m_debugApplication64;
    internal readonly IDebugApplication32 m_debugApplication32;

    public int GetDocVersionString(out string version)
    {
        version = new Version().ToString();
        System.Diagnostics.Debug.WriteLine($"{nameof(ScriptSite)}.{nameof(GetDocVersionString)} {version}");
        return S_OK;
    }

    internal Dictionary<string, object> NamedItems { get; } = new Dictionary<string, object>();

    const uint TYPE_E_ELEMENTNOTFOUND = 0x8002802B;
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
        System.Diagnostics.Debug.WriteLine($"{nameof(ScriptSite)}.{nameof(OnLeaveScript)}");
        return S_OK;
    }

    public int OnScriptError(IActiveScriptError scriptError)
    {
        System.Diagnostics.Debug.WriteLine($"{nameof(ScriptSite)}.{nameof(OnScriptError)} {scriptError}");
        return 0;
    }

    public int OnScriptTerminate(ref object result, stdole.EXCEPINFO[] exceptionInfo)
    {
        System.Diagnostics.Debug.WriteLine($"{nameof(ScriptSite)}.{nameof(OnScriptTerminate)} {result}");
        return S_OK;
    }

    public int OnStateChange(SCRIPTSTATE scriptState)
    {
        System.Diagnostics.Debug.WriteLine($"{nameof(ScriptSite)}.{nameof(OnStateChange)} {scriptState}");
        return S_OK;
    }
    int IActiveScriptSiteDebug64.GetDocumentContextFromPosition(ulong dwSourceContext, uint uCharacterOffset, uint uNumChars, out IDebugDocumentContext ppsc)
    {
        System.Diagnostics.Debug.WriteLine($"{nameof(ScriptSite)}64.{nameof(IActiveScriptSiteDebug64.GetDocumentContextFromPosition)} {dwSourceContext}, {uCharacterOffset} {uNumChars}");
        throw new NotImplementedException();
    }
    int IActiveScriptSiteDebug32.GetDocumentContextFromPosition(uint dwSourceContext, uint uCharacterOffset, uint uNumChars, out IDebugDocumentContext ppsc)
    {
        System.Diagnostics.Debug.WriteLine($"{nameof(ScriptSite)}32.{nameof(IActiveScriptSiteDebug32.GetDocumentContextFromPosition)} {dwSourceContext}, {uCharacterOffset} {uNumChars}");
        throw new NotImplementedException();
    }

    int IActiveScriptSiteDebug32.GetApplication(out IDebugApplication32 ppda)
    {
        System.Diagnostics.Debug.WriteLine($"{nameof(ScriptSite)}32.{nameof(IActiveScriptSiteDebug32.GetApplication)}");
        ppda = m_debugApplication32;
        return S_OK;
    }
    int IActiveScriptSiteDebug64.GetApplication(out IDebugApplication64 ppda)
    {
        System.Diagnostics.Debug.WriteLine($"{nameof(ScriptSite)}64.{nameof(IActiveScriptSiteDebug64.GetApplication)}");
        ppda = m_debugApplication64;
        return S_OK;
    }

    int IActiveScriptSiteDebug64.GetRootApplicationNode(out IDebugApplicationNode ppdanRoot)
    {
        SUCCESS(m_debugApplication64.GetRootNode(out ppdanRoot));
        SUCCESS(ppdanRoot.GetName(DOCUMENTNAMETYPE.DOCUMENTNAMETYPE_TITLE, out var title));
        System.Diagnostics.Debug.WriteLine($"{nameof(ScriptSite)}64.{nameof(IActiveScriptSiteDebug64.GetRootApplicationNode)} {title}");
        return 0x8004001;
    }
    int IActiveScriptSiteDebug32.GetRootApplicationNode(out IDebugApplicationNode ppdanRoot)
    {
        SUCCESS(m_debugApplication32.GetRootNode(out ppdanRoot));
        SUCCESS(ppdanRoot.GetName(DOCUMENTNAMETYPE.DOCUMENTNAMETYPE_TITLE, out var title));
        System.Diagnostics.Debug.WriteLine($"{nameof(ScriptSite)}32.{nameof(IActiveScriptSiteDebug32.GetRootApplicationNode)} {title}");
        return 0x8004001;
    }

    public int OnScriptErrorDebug(IActiveScriptErrorDebug pErrorDebug, out int pfEnterDebugger, out int pfCallOnScriptErrorWhenContinuing)
    {
        System.Diagnostics.Debug.WriteLine($"{nameof(ScriptSite)}.{nameof(OnScriptErrorDebug)} {pErrorDebug}");
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
