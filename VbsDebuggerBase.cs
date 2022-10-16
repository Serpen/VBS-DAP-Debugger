using Microsoft.VisualStudio.Debugger.Interop;
using ActiveDbg;
using static Helpers;
using Marshal = System.Runtime.InteropServices.Marshal;
using stdole;

public class VbsDebuggerBase : IDisposable
{
#if ARCH64
    static internal IProcessDebugManager64 pdm;

    internal IDebugApplication64 debugApplication;
#else
    static internal IProcessDebugManager32 pdm;
    internal IDebugApplication32 debugApplication;
#endif

    private IApplicationDebugger applicationDebugger;
    static IActiveScriptMy languageEngine;
    private IActiveScriptSiteMy scriptSite;
    internal IRemoteDebugApplicationThread DebugThread;

    readonly static Type PDMtype = Type.GetTypeFromProgID("ProcessDebugManager") ?? throw new Exception("no def ProcessDebugManager");
    readonly static Type VBScriptType = System.Type.GetTypeFromProgID("VBScript") ?? throw new Exception("no def VBScript");
    static readonly Type MSProgramProvider2Type = Type.GetTypeFromCLSID(new Guid("{170EC3FC-4E80-40AB-A85A-55900C7C70DE}")) ?? throw new Exception("no pdm2type");

    static readonly CONST_GUID_ARRAY ScriptEngineFilter;

    static VbsDebuggerBase()
    {
#if ARCH64
        pdm = Activator.CreateInstance(PDMtype) as IProcessDebugManager64 ?? throw new Exception($"no {nameof(IProcessDebugManager64)}");
#else
        //pdm = Activator.CreateInstance(PDMtype) as IProcessDebugManager32 ?? throw new Exception($"no {nameof(IProcessDebugManager32)}");
#endif
        languageEngine = Activator.CreateInstance(VBScriptType) as IActiveScriptMy ?? throw new Exception($"no {nameof(IActiveScriptMy)}");

        var sefPtr = Marshal.AllocHGlobal((Marshal.SizeOf<Guid>()));

        Guid sefguid = new Guid("{F200A7E7-DEA5-11D0-B854-00A0244A1DE2}");
        Marshal.StructureToPtr(sefguid, sefPtr, false);

        ScriptEngineFilter = new CONST_GUID_ARRAY() { dwCount = 1, Members = sefPtr };
    }
    public VbsDebuggerBase()
    {
        //SUCCESS(m_processDebugManager.GetDefaultApplication(out var m_debugApplication));
        SUCCESS(pdm.CreateApplication(out debugApplication));
        ArgumentNullException.ThrowIfNull(debugApplication);

        SUCCESS(debugApplication.SetName(nameof(VbsDebuggerBase)));

        SUCCESS(pdm.AddApplication(debugApplication, out var cookie));

        applicationDebugger = new Debugger();

        var debugSessionProvider = applicationDebugger as IDebugSessionProvider ?? throw new Exception($"no {nameof(IDebugSessionProvider)}");

        SUCCESS(debugSessionProvider.StartDebugSession(debugApplication));

        scriptSite = new ScriptSite(debugApplication) ?? throw new Exception($"no {nameof(IActiveScriptSite)}");

        // border for -> Parse
    }

    internal void setBreakPoint(uint line)
    {
        var iasd64 = languageEngine as IActiveScriptDebug64 ?? throw new Exception("no IActiveScriptDebug64");

        SUCCESS(iasd64.EnumCodeContextsOfPosition(0, line, 0, out var enumDebugCodeContexts));

        SUCCESS(enumDebugCodeContexts.Next(1, out var dcc, out var _));

        SUCCESS(dcc.SetBreakPoint(BREAKPOINT_STATE.BREAKPOINT_ENABLED));
    }

    internal void Resume(BREAKRESUMEACTION resumeAction = BREAKRESUMEACTION.BREAKRESUMEACTION_CONTINUE)
    {
        if (DebugThread is not null)
        {
            DebugThread.GetApplication(out var rdApp);
            rdApp.ResumeFromBreakPoint(DebugThread, resumeAction, ERRORRESUMEACTION.ERRORRESUMEACTION_SkipErrorStatement);
        }
    }

    public void Parse(string scriptText)
    {
        // border for <- .ctor
        SUCCESS(languageEngine.SetScriptSite(scriptSite));
#if ARCH64
        var _parse = languageEngine as IActiveScriptParse64 ?? throw new Exception($"no {nameof(IActiveScriptParse64)}");
#else
        var _parse = languageEngine as IActiveScriptParse32 ?? throw new Exception($"no {nameof(IActiveScriptParse32)}");
#endif
        SUCCESS(_parse.InitNew());

        TestClass myObj = new TestClass("Hallo", 1);
        SUCCESS(languageEngine.AddNamedItem(nameof(myObj), (uint)(ScriptItem.IsVisible | ScriptItem.IsSource)));
        (scriptSite as ScriptSite).NamedItems.Add(nameof(myObj), myObj);

        var obj = new stdole.EXCEPINFO[1];
        SUCCESS(_parse.ParseScriptText(scriptText, null, null, null, 0, 0, (uint)(ScriptText.IsVisible), out var result, null), throwException: true);
        System.Console.WriteLine("ParseScriptText finished " + myObj.Name);
    }

    public object Invoke(string expression)
    {
        // border for <- .ctor
        SUCCESS(languageEngine.SetScriptSite(scriptSite));
#if ARCH64
        var _parse = languageEngine as IActiveScriptParse64 ?? throw new Exception($"no {nameof(IActiveScriptParse64)}");
#else
        var _parse = languageEngine as IActiveScriptParse32 ?? throw new Exception($"no {nameof(IActiveScriptParse32)}");
#endif
        SUCCESS(_parse.InitNew());

        SUCCESS(_parse.ParseScriptText(expression, null, null, null, 0, 0, (uint)(ScriptText.IsVisible | ScriptText.IsExpression), out var result, null), throwException: true);
        System.Console.WriteLine("ParseScriptText finished " + result);
        return result;
    }

    public void Dispose()
    {
        SUCCESS(debugApplication.DisconnectDebugger());
        SUCCESS(languageEngine.Close());
    }

    internal static IEnumerable<System.Diagnostics.Process> GetScriptProcesses()
    {
        var outlist = new List<System.Diagnostics.Process>();
        var pdm2 = Activator.CreateInstance(MSProgramProvider2Type) as IDebugProgramProvider2My ?? throw new Exception("no IDebugProgramProvider2");

        foreach (var proc in System.Diagnostics.Process.GetProcesses())
        // foreach (var proc in System.Diagnostics.Process.GetProcessesByName("iexplore"))
        {
            System.Diagnostics.Debug.WriteLine("process {0} {1}...", proc, proc.Id);

            var adprocid = new AD_PROCESS_ID() { dwProcessId = (uint)proc.Id, ProcessIdType = (uint)enum_AD_PROCESS_ID.AD_PROCESS_ID_SYSTEM };

            int result = 1;
            try
            {
                result = pdm2.GetProviderProcessData((enum_PROVIDER_FLAGS.PFLAG_GET_PROGRAM_NODES), null, adprocid, ScriptEngineFilter, out var provdata);

                SUCCESS(result);

                if (result == 0 && provdata.ProgramNodes.dwCount > 0)
                    outlist.Add(proc);
            }
            catch { }
        }
        return outlist;
    }


    // not working
    internal static void Attach(System.Diagnostics.Process proc)
    {
        var outlist = new List<System.Diagnostics.Process>();
        var pdm2 = Activator.CreateInstance(MSProgramProvider2Type) as IDebugProgramProvider2My ?? throw new Exception("no IDebugProgramProvider2");

        System.Diagnostics.Debug.WriteLine("process {0} {1}...", proc, proc.Id);

        var adprocid = new AD_PROCESS_ID() { dwProcessId = (uint)proc.Id, ProcessIdType = (uint)enum_AD_PROCESS_ID.AD_PROCESS_ID_SYSTEM };

        int result = 1;
        try
        {
            result = pdm2.GetProviderProcessData(enum_PROVIDER_FLAGS.PFLAG_GET_PROGRAM_NODES | enum_PROVIDER_FLAGS.PFLAG_DEBUGGEE | enum_PROVIDER_FLAGS.PFLAG_ATTACHED_TO_DEBUGGEE, null, adprocid, ScriptEngineFilter, out var provdata);

            SUCCESS(result);

            var dpn2Guid = new Guid("426E255C-F1CE-4D02-A931-F9A254BF7F0F");

            Marshal.QueryInterface(provdata.ProgramNodes.Members, ref dpn2Guid, out var ptr1);

            var dpn2_ = Marshal.GetObjectForIUnknown(ptr1);
            var dpn2 = dpn2_ as IDebugProgramNode2;
            var dppn2 = dpn2 as IDebugProviderProgramNode2;

            var rdaGuid = new Guid("51973C30-CB0C-11D0-B5C9-00A0244A0E7A");
            dppn2.UnmarshalDebuggeeInterface(ref rdaGuid, out var ptr2);

            var rda = Marshal.GetObjectForIUnknown(ptr2) as IRemoteDebugApplication;

            //  dpn2 = provdata.ProgramNodes.Members as IDebugProgramNode2;
        }
        catch { }
    }

    internal void Wait(System.Diagnostics.Process proc)
    {
        var pdm2_ = Activator.CreateInstance(MSProgramProvider2Type) ?? throw new Exception("no pdm2_");
        var pdm2 = pdm2_ as ActiveDbg.IDebugProgramProvider2My ?? throw new Exception("no IDebugProgramProvider2"); ;

        var adprocid = new AD_PROCESS_ID() { dwProcessId = (uint)proc.Id, ProcessIdType = (uint)enum_AD_PROCESS_ID.AD_PROCESS_ID_SYSTEM };

        var septr = Marshal.AllocHGlobal((Marshal.SizeOf<Guid>()));
        Marshal.StructureToPtr(ScriptEngineFilter, septr, false);

        var _callback = new Callback();
        CONST_GUID_ARRAY sef = new CONST_GUID_ARRAY() { dwCount = 1, Members = septr }; // };

        SUCCESS(pdm2.WatchForProviderEvents(
            enum_PROVIDER_FLAGS.PFLAG_DEBUGGEE | enum_PROVIDER_FLAGS.PFLAG_REASON_WATCH,
            null,
            adprocid,
            sef,
            Guid.Empty,
            _callback
            ));

        _callback.Wait();

        // unreg
        SUCCESS(pdm2.WatchForProviderEvents(
            enum_PROVIDER_FLAGS.PFLAG_NONE,
            null,
            adprocid,
            sef,
            Guid.Empty,
            _callback
            ));


    }

}