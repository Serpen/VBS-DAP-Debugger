// using ActiveDbg;
using Microsoft.VisualStudio.Debugger.Interop;

#if ARCH64
    class ScriptSite : ActiveDbg.IActiveScriptSiteMy, IActiveScriptSiteDebug64, IActiveScriptSiteWindow
    {
        public ScriptSite(IDebugApplication64 debugApplication)
#else
    class CMyScriptSite : IActiveScriptSite, IActiveScriptSiteDebug32, IActiveScriptSiteWindow
    {
    public CMyScriptSite(IDebugApplication32 debugApplication)
#endif
        {
            m_debugApplication = debugApplication;
        }

#if ARCH64
        internal IDebugApplication64 m_debugApplication;
#else
        internal IDebugApplication32 m_debugApplication;
#endif
        public int GetDocVersionString(out string version)
        {
            version = new Version().ToString();
            System.Diagnostics.Debug.WriteLine("MyScriptSite.GetDocVersionString " + version);
            return 0;
        }

        internal Dictionary<string, object> NamedItems { get; } = new Dictionary<string, object>();

        const uint TYPE_E_ELEMENTNOTFOUND = 0x8002802B;
        public uint GetItemInfo(string pstrName, uint dwReturnMask, out object ppiunkItem, IntPtr ppti)
        {
            System.Diagnostics.Debug.WriteLine("MyScriptSite.GetItemInfo {0} {1}", pstrName, dwReturnMask);
            if (NamedItems.TryGetValue(pstrName, out var item))
            {
                ppiunkItem = item;
                return 0; 
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
            System.Diagnostics.Debug.WriteLine("MyScriptSite.GetLCID " + lcid);
            return 0;
        }

        public int OnEnterScript()
        {
            System.Diagnostics.Debug.WriteLine("MyScriptSite.OnEnterScript");
            return 0;
        }

        public int OnLeaveScript()
        {
            System.Diagnostics.Debug.WriteLine("MyScriptSite.OnLeaveScript");
            return 0;
        }

        public int OnScriptError(IActiveScriptError scriptError)
        {
            System.Diagnostics.Debug.WriteLine("MyScriptSite.OnScriptError {0}", scriptError);
            return 0;
        }

        public int OnScriptTerminate(ref object result, stdole.EXCEPINFO[] exceptionInfo)
        {
            System.Diagnostics.Debug.WriteLine("MyScriptSite.OnScriptTerminate {0}", result);
            return 0;
        }

        public int OnStateChange(SCRIPTSTATE scriptState)
        {
            System.Diagnostics.Debug.WriteLine("MyScriptSite.OnStateChange {0}", scriptState);
            return 0;
        }

#if ARCH64
        public int GetDocumentContextFromPosition(ulong dwSourceContext, uint uCharacterOffset, uint uNumChars, out IDebugDocumentContext ppsc)
#else
        public int GetDocumentContextFromPosition(uint dwSourceContext, uint uCharacterOffset, uint uNumChars, out IDebugDocumentContext ppsc)

#endif
        {
            System.Diagnostics.Debug.WriteLine("MyScriptSite.GetDocumentContextFromPosition {0}, {1} {2}", dwSourceContext, uCharacterOffset, uNumChars);
            throw new NotImplementedException();
        }

#if ARCH64
        public int GetApplication(out IDebugApplication64 ppda)
#else
        public int GetApplication(out IDebugApplication32 ppda)
#endif

        {
            ppda = m_debugApplication;
            return 0;
        }

        public int GetRootApplicationNode(out IDebugApplicationNode ppdanRoot)
        {
            m_debugApplication.GetRootNode(out ppdanRoot);
            System.Diagnostics.Debug.WriteLine("MyScriptSite.GetRootApplicationNode {0}", ppdanRoot);
            return 0x8004001;
        }

        public int OnScriptErrorDebug(IActiveScriptErrorDebug pErrorDebug, out int pfEnterDebugger, out int pfCallOnScriptErrorWhenContinuing)
        {
            pfEnterDebugger = 0;
            pfCallOnScriptErrorWhenContinuing = 1;
            System.Diagnostics.Debug.WriteLine("MyScriptSite.OnScriptErrorDebug {0}", pErrorDebug);
            return 0;
        }

        public int GetWindow(out IntPtr phwnd)
        {
            phwnd = IntPtr.Zero;
            System.Diagnostics.Debug.WriteLine("MyScriptSite.GetWindow {0}", phwnd);
            return 0;
        }

        public int EnableModeless(int fEnable)
        {
            System.Diagnostics.Debug.WriteLine("MyScriptSite.EnableModeless {0}", fEnable);
            return 0;
        }

        
    }
