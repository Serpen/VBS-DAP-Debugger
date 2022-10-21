using Microsoft.VisualStudio.Debugger.Interop;
using ActiveDbg;
using static Helpers;
using Marshal = System.Runtime.InteropServices.Marshal;

public class VbsDebuggerBase : IDisposable
{

    static internal IProcessDebugManager64 pdm64;
    internal IDebugApplication64 debugApplication64;

    static internal IProcessDebugManager32 pdm32;
    internal IDebugApplication32 debugApplication32;


    private Debugger applicationDebugger;
    static IActiveScriptMy languageEngine;
    private IActiveScriptSiteMy scriptSite;
    internal IRemoteDebugApplicationThread DebugThread;

    readonly static Type PDMtype = Type.GetTypeFromProgID("ProcessDebugManager") ?? throw new Exception("no def ProcessDebugManager");
    readonly static Type VBScriptType = System.Type.GetTypeFromProgID("VBScript") ?? throw new Exception("no def VBScript");
    static readonly Type MSProgramProvider2Type = Type.GetTypeFromCLSID(new Guid("{170EC3FC-4E80-40AB-A85A-55900C7C70DE}")) ?? throw new Exception("no pdm2type");

    static readonly CONST_GUID_ARRAY ScriptEngineFilter;

    static VbsDebuggerBase()
    {
        pdm64 = Activator.CreateInstance(PDMtype) as IProcessDebugManager64; // ?? throw new Exception($"no {nameof(IProcessDebugManager64)}");
        pdm32 = Activator.CreateInstance(PDMtype) as IProcessDebugManager32; // ?? throw new Exception($"no {nameof(IProcessDebugManager32)}");

        languageEngine = Activator.CreateInstance(VBScriptType) as IActiveScriptMy ?? throw new Exception($"no {nameof(IActiveScriptMy)}");

        var sefPtr = Marshal.AllocHGlobal((Marshal.SizeOf<Guid>()));

        Guid sefguid = new Guid("{F200A7E7-DEA5-11D0-B854-00A0244A1DE2}");
        Marshal.StructureToPtr(sefguid, sefPtr, false);

        ScriptEngineFilter = new CONST_GUID_ARRAY() { dwCount = 1, Members = sefPtr };
    }
    public VbsDebuggerBase()
    {
        //SUCCESS(m_processDebugManager.GetDefaultApplication(out var m_debugApplication));
        if (pdm64 is not null)
        {
            SUCCESS(pdm64.CreateApplication(out debugApplication64));
            ArgumentNullException.ThrowIfNull(debugApplication64);
            SUCCESS(debugApplication64.SetName(nameof(VbsDebuggerBase)));
            SUCCESS(pdm64.AddApplication(debugApplication64, out var cookie));
            applicationDebugger = new Debugger();
            var debugSessionProvider = applicationDebugger as IDebugSessionProvider ?? throw new Exception($"no {nameof(IDebugSessionProvider)}");
            SUCCESS(debugSessionProvider.StartDebugSession(debugApplication64));
            scriptSite = new ScriptSite(debugApplication64) ?? throw new Exception($"no {nameof(IActiveScriptSite)}");
        }
        else
        {
            SUCCESS(pdm32.CreateApplication(out debugApplication32));
            ArgumentNullException.ThrowIfNull(debugApplication32);
            SUCCESS(debugApplication32.SetName(nameof(VbsDebuggerBase)));
            SUCCESS(pdm32.AddApplication(debugApplication32, out var cookie));
            applicationDebugger = new Debugger();
            var debugSessionProvider = applicationDebugger as IDebugSessionProvider ?? throw new Exception($"no {nameof(IDebugSessionProvider)}");
            SUCCESS(debugSessionProvider.StartDebugSession(debugApplication32));
            scriptSite = new ScriptSite(debugApplication32) ?? throw new Exception($"no {nameof(IActiveScriptSite)}");
        }

        // border for -> Parse
    }

    internal void CloseEvent()
    {
        connected = false;
    }


    internal void setBreakPoint(uint line)
    {
        var iasd32 = languageEngine as IActiveScriptDebug32;
        var iasd64 = languageEngine as IActiveScriptDebug64;
        IDebugCodeContext dcc;
        IEnumDebugCodeContexts enumDebugCodeContexts;

        if (iasd64 is not null)
            SUCCESS(iasd64.EnumCodeContextsOfPosition(0, line, 0, out enumDebugCodeContexts));
        else
            SUCCESS(iasd32.EnumCodeContextsOfPosition(0, line, 0, out enumDebugCodeContexts));

        SUCCESS(enumDebugCodeContexts.Next(1, out dcc, out var _));

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
        SUCCESS(languageEngine.SetScriptSite(scriptSite));

        var parser32 = languageEngine as IActiveScriptParse32;
        var parser64 = languageEngine as IActiveScriptParse64;

        if (parser64 is not null)
            Parse64(parser64, scriptText, ScriptText.IsVisible);
        else
            Parse32(parser32, scriptText, ScriptText.IsVisible);
    }

    public object Invoke(string scriptText)
    {
        SUCCESS(languageEngine.SetScriptSite(scriptSite));

        var parser32 = languageEngine as IActiveScriptParse32;
        var parser64 = languageEngine as IActiveScriptParse64;

        if (parser64 is not null)
            return Parse64(parser64, scriptText, ScriptText.IsVisible | ScriptText.IsExpression);
        else
            return Parse32(parser32, scriptText, ScriptText.IsVisible | ScriptText.IsExpression);
    }


    object Parse32(IActiveScriptParse32 parser, string scriptText, ScriptText flags)
    {
        // border for <- .ctor
        SUCCESS(parser.InitNew());

        // TestClass myObj = new TestClass("Hallo", 1);
        // SUCCESS(languageEngine.AddNamedItem(nameof(myObj), (uint)(ScriptItem.IsVisible | ScriptItem.IsSource)));
        // (scriptSite as ScriptSite).NamedItems.Add(nameof(myObj), myObj);

        var obj = new stdole.EXCEPINFO[1];
        SUCCESS(parser.ParseScriptText(scriptText, null, null, null, 0, 0, (uint)flags, out var result, null), throwException: true);
        // System.Console.WriteLine("ParseScriptText finished " + myObj.Name);
        return result;
    }


    object Parse64(IActiveScriptParse64 parser, string scriptText, ScriptText flags)
    {
        // border for <- .ctor
        SUCCESS(parser.InitNew());

        // TestClass myObj = new TestClass("Hallo", 1);
        // SUCCESS(languageEngine.AddNamedItem(nameof(myObj), (uint)(ScriptItem.IsVisible | ScriptItem.IsSource)));
        // (scriptSite as ScriptSite).NamedItems.Add(nameof(myObj), myObj);

        var obj = new stdole.EXCEPINFO[1];
        SUCCESS(parser.ParseScriptText(scriptText, null, null, null, 0, 0, (uint)flags, out var result, null), throwException: true);
        // System.Console.WriteLine("ParseScriptText finished " + myObj.Name);
        return result;
    }

    public void Dispose()
    {
        debugApplication32?.DisconnectDebugger();
        debugApplication64?.DisconnectDebugger();
        SUCCESS(languageEngine.Close());
    }

    internal static IEnumerable<System.Diagnostics.Process> GetScriptProcesses()
    {
        var outlist = new List<System.Diagnostics.Process>();
        var pdm2 = Activator.CreateInstance(MSProgramProvider2Type) as IDebugProgramProvider2My ?? throw new Exception("no IDebugProgramProvider2");

        // foreach (var proc in System.Diagnostics.Process.GetProcesses())
        foreach (var proc in System.Diagnostics.Process.GetProcessesByName("iexplore"))
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

    bool connected;

    internal void Attach(System.Diagnostics.Process proc)
    {

        var pdm2 = Activator.CreateInstance(MSProgramProvider2Type) as IDebugProgramProvider2My ?? throw new Exception("no IDebugProgramProvider2");

        var adprocid = new AD_PROCESS_ID() { dwProcessId = (uint)proc.Id, ProcessIdType = (uint)enum_AD_PROCESS_ID.AD_PROCESS_ID_SYSTEM };

        int result = 1;
        try
        {
            result = pdm2.GetProviderProcessData(enum_PROVIDER_FLAGS.PFLAG_GET_PROGRAM_NODES | enum_PROVIDER_FLAGS.PFLAG_DEBUGGEE | enum_PROVIDER_FLAGS.PFLAG_ATTACHED_TO_DEBUGGEE, null, adprocid, ScriptEngineFilter, out var provdata);

            SUCCESS(result);


            var dpn2Guid = new Guid((typeof(IDebugProgramNode2).GetCustomAttributes(typeof(System.Runtime.InteropServices.GuidAttribute), false).First() as System.Runtime.InteropServices.GuidAttribute).Value); // new Guid("426E255C-F1CE-4D02-A931-F9A254BF7F0F");
            var rdaGuid = new Guid((typeof(IRemoteDebugApplication).GetCustomAttributes(typeof(System.Runtime.InteropServices.GuidAttribute), false).First() as System.Runtime.InteropServices.GuidAttribute).Value); // new Guid("51973C30-CB0C-11D0-B5C9-00A0244A0E7A"); 

            var ptrptr = Marshal.ReadIntPtr(provdata.ProgramNodes.Members);


            Marshal.QueryInterface(ptrptr, ref dpn2Guid, out var ptr1);

            var dpn2_ = Marshal.GetObjectForIUnknown(ptr1);
            // var dpn2 = dpn2_ as IDebugProgramNode2;
            var dppn2 = dpn2_ as IDebugProviderProgramNode2;

            dppn2.UnmarshalDebuggeeInterface(ref rdaGuid, out var ptr2);

            var rda = Marshal.GetObjectForIUnknown(ptr2) as IRemoteDebugApplication;

            System.Diagnostics.Debug.WriteLine("process {0} {1} {3}", proc, proc.Id, rda.GetName(out var rdaname), rdaname);

            rda.ConnectDebugger(applicationDebugger);
            // rda.GetRootNode(out var dan);
            // dan.GetName(DOCUMENTNAMETYPE.DOCUMENTNAMETYPE_TITLE, out var name1);
            
            //rda.CauseBreak();
            connected = true;

            // while (connected) {
            //     System.Threading.Thread.Sleep(10000);
            // }

            // var debugSessionProvider = applicationDebugger as IDebugSessionProvider ?? throw new Exception($"no {nameof(IDebugSessionProvider)}");

            // SUCCESS(debugSessionProvider.StartDebugSession(debugApplication));

            // scriptSite = new ScriptSite(debugApplication) ?? throw new Exception($"no {nameof(IActiveScriptSite)}");


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