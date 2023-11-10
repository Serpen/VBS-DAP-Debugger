using Microsoft.VisualStudio.Debugger.Interop;
using ActiveDbg;
using static Helpers;
using Marshal = System.Runtime.InteropServices.Marshal;
using DAP = Microsoft.VisualStudio.Shared.VSCodeDebugProtocol;
using Microsoft.VisualStudio.Shared.VSCodeDebugProtocol.Messages;
using System.Diagnostics;

public class VbsDebugAdapter : DAP.DebugAdapterBase, IDisposable
{

    readonly static IProcessDebugManager64 pdm64;
    readonly IDebugApplication64 debugApplication64;

    readonly static IProcessDebugManager32 pdm32;
    readonly IDebugApplication32 debugApplication32;

    internal readonly Debugger applicationDebugger;

    readonly static dynamic VBScript;
    readonly static IActiveScriptMy languageEngine;
    readonly IActiveScriptSiteMy scriptSite;
    internal IRemoteDebugApplicationThread DebugThread;

    static readonly Type MSProgramProvider2Type = Type.GetTypeFromCLSID(new Guid("{170EC3FC-4E80-40AB-A85A-55900C7C70DE}")); // ?? throw new Exception("no pdm2type");
    static readonly IDebugProgramProvider2My64 MSProgramProvider2;

    static readonly CONST_GUID_ARRAY ScriptEngineFilter;

    readonly object syncobject = new object();

    bool parsedInited = true;

    static VbsDebugAdapter()
    {
        var PDMtype = Type.GetTypeFromProgID("ProcessDebugManager") ?? throw new Exception("no def ProcessDebugManager");
        var VBScriptType = Type.GetTypeFromProgID("VBScript") ?? throw new Exception("no def VBScript");

        pdm64 = Activator.CreateInstance(PDMtype) as IProcessDebugManager64; // ?? throw new Exception($"no {nameof(IProcessDebugManager64)}");
        pdm32 = Activator.CreateInstance(PDMtype) as IProcessDebugManager32; // ?? throw new Exception($"no {nameof(IProcessDebugManager32)}");

        VBScript = Activator.CreateInstance(VBScriptType) ?? throw new Exception($"no {nameof(VBScript)}");
        languageEngine = VBScript as IActiveScriptMy ?? throw new Exception($"no {nameof(IActiveScriptMy)}");

        var sefPtr = Marshal.AllocHGlobal((Marshal.SizeOf<Guid>()));

        Guid sefguid = new Guid("{F200A7E7-DEA5-11D0-B854-00A0244A1DE2}");
        Marshal.StructureToPtr(sefguid, sefPtr, false);

        ScriptEngineFilter = new CONST_GUID_ARRAY() { dwCount = 1, Members = sefPtr };

        MSProgramProvider2 = Activator.CreateInstance(MSProgramProvider2Type) as IDebugProgramProvider2My64;
    }
    public VbsDebugAdapter()
    {
        DebugWriteMethodeName();
        //SUCCESS(m_processDebugManager.GetDefaultApplication(out var m_debugApplication));
        if (pdm64 is not null)
        {
            SUCCESS(pdm64.CreateApplication(out debugApplication64));
            ArgumentNullException.ThrowIfNull(debugApplication64);
            SUCCESS(debugApplication64.SetName(nameof(VbsDebugAdapter)));
            SUCCESS(pdm64.AddApplication(debugApplication64, out var cookie));
            applicationDebugger = new Debugger(this);
            var debugSessionProvider = applicationDebugger as IDebugSessionProvider ?? throw new Exception($"no {nameof(IDebugSessionProvider)}");
            SUCCESS(debugSessionProvider.StartDebugSession(debugApplication64));
            scriptSite = new ScriptSite(debugApplication64, this) ?? throw new Exception($"no {nameof(IActiveScriptSite)}");
        }
        else
        {
            SUCCESS(pdm32.CreateApplication(out debugApplication32));
            ArgumentNullException.ThrowIfNull(debugApplication32);
            SUCCESS(debugApplication32.SetName(nameof(VbsDebugAdapter)));
            SUCCESS(pdm32.AddApplication(debugApplication32, out var cookie));
            applicationDebugger = new Debugger(this);
            var debugSessionProvider = applicationDebugger as IDebugSessionProvider ?? throw new Exception($"no {nameof(IDebugSessionProvider)}");
            SUCCESS(debugSessionProvider.StartDebugSession(debugApplication32));
            scriptSite = new ScriptSite(debugApplication32, this) ?? throw new Exception($"no {nameof(IActiveScriptSite)}");
        }

        System.Diagnostics.Debug.WriteLine("DAP.VbsDebuggerBase InitializeProtocolClient");

        InitializeProtocolClient(System.Console.OpenStandardInput(), System.Console.OpenStandardOutput());

        System.Diagnostics.Debug.WriteLine("DAP.VbsDebuggerBase end");

        // border for -> Parse
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
        lock (syncobject)
        {
            if (DebugThread is not null)
            {
                DebugThread.GetApplication(out var rdApp);
                rdApp?.ResumeFromBreakPoint(DebugThread, resumeAction, ERRORRESUMEACTION.ERRORRESUMEACTION_SkipErrorStatement);
            }
        }
    }

    public void Parse(string scriptText, IDictionary<string, object> namedItems = null)
    {
        var parser32 = languageEngine as IActiveScriptParse32;
        var parser64 = languageEngine as IActiveScriptParse64;

        if (parser64 is not null)
            Parse64(parser64, scriptText, ScriptText.IsVisible, namedItems);
        else
            Parse32(parser32, scriptText, ScriptText.IsVisible, namedItems);
    }

    public object Invoke(string scriptText, IDictionary<string, object> namedItems = null)
    {
        var parser32 = languageEngine as IActiveScriptParse32;
        var parser64 = languageEngine as IActiveScriptParse64;

        if (parser64 is not null)
            return Parse64(parser64, scriptText, ScriptText.IsVisible | ScriptText.IsExpression, namedItems);
        else
            return Parse32(parser32, scriptText, ScriptText.IsVisible | ScriptText.IsExpression, namedItems);
    }


    object Parse32(IActiveScriptParse32 parser, string scriptText, ScriptText flags, IDictionary<string, object> namedItems)
    {
        // border for <- .ctor

        // must be set on same thread??
        SUCCESS(languageEngine.SetScriptSite(scriptSite));

        if (!parsedInited)
        {
            SUCCESS(parser.InitNew());
            parsedInited = true;
        }

        namedItems ??= new Dictionary<string, object>();

        foreach (var obj in namedItems)
        {
            SUCCESS(languageEngine.AddNamedItem(obj.Key, (uint)(ScriptItem.IsVisible | ScriptItem.IsSource)));
            (scriptSite as ScriptSite).NamedItems.Add(obj.Key, obj.Value);
        }

        SUCCESS(parser.ParseScriptText(scriptText, null, null, null, 0, 0, (uint)flags, out var result, null), throwException: true);
        return result;
    }


    object Parse64(IActiveScriptParse64 parser, string scriptText, ScriptText flags, IDictionary<string, object> namedItems)
    {
        // border for <- .ctor

        // must be set on same thread??
        SUCCESS(languageEngine.SetScriptSite(scriptSite));

        if (!parsedInited)
        {
            SUCCESS(parser.InitNew());
            parsedInited = true;
        }

        namedItems ??= new Dictionary<string, object>();

        foreach (var obj in namedItems)
        {
            SUCCESS(languageEngine.AddNamedItem(obj.Key, (uint)(ScriptItem.IsVisible | ScriptItem.IsSource)));
            (scriptSite as ScriptSite).NamedItems.Add(obj.Key, obj.Value);
        }
        object result = null;
        try
        {
            SUCCESS(parser.ParseScriptText(scriptText, null, null, null, 0, 0, (uint)flags, out result, null), throwException: true);
        }
        catch { }
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

        foreach (var proc in System.Diagnostics.Process.GetProcesses().Where(p => p.ProcessName.Contains("script", StringComparison.CurrentCultureIgnoreCase)))
        //foreach (var proc in System.Diagnostics.Process.GetProcessesByName("cscript"))
        {
            if (proc.Id == Process.GetCurrentProcess().Id)
                continue;
            System.Diagnostics.Debug.WriteLine("process {0} {1}...", proc, proc.Id);

            var adprocid = new AD_PROCESS_ID() { dwProcessId = (uint)proc.Id, ProcessIdType = (uint)enum_AD_PROCESS_ID.AD_PROCESS_ID_SYSTEM };

            int result = 1;
            try
            {
                result = MSProgramProvider2.GetProviderProcessData(enum_PROVIDER_FLAGS.PFLAG_GET_PROGRAM_NODES, null, adprocid, ScriptEngineFilter, out var provdata);

                SUCCESS(result);

                if (result == 0 && provdata.ProgramNodes.dwCount > 0)
                    outlist.Add(proc);
            }
            catch { }
        }
        return outlist;
    }

    internal void Attach(System.Diagnostics.Process proc)
    {
        var adprocid = new AD_PROCESS_ID() { dwProcessId = (uint)proc.Id, ProcessIdType = (uint)enum_AD_PROCESS_ID.AD_PROCESS_ID_SYSTEM };

        int result = 1;
        try
        {
            result = MSProgramProvider2.GetProviderProcessData(
                enum_PROVIDER_FLAGS.PFLAG_GET_PROGRAM_NODES | enum_PROVIDER_FLAGS.PFLAG_DEBUGGEE | enum_PROVIDER_FLAGS.PFLAG_ATTACHED_TO_DEBUGGEE,
                null, adprocid, ScriptEngineFilter, out var provdata);
            

            SUCCESS(result);

            var dpn2Guid = GetInterfaceGuid(typeof(IDebugProgramNode2));
            var rdaGuid = GetInterfaceGuid(typeof(IRemoteDebugApplication));

            var ptrptr = Marshal.ReadIntPtr(provdata.ProgramNodes.Members);

            Marshal.QueryInterface(ptrptr, ref dpn2Guid, out var ptr1);

            var dppn2 = Marshal.GetObjectForIUnknown(ptr1) as IDebugProviderProgramNode2;

            dppn2.UnmarshalDebuggeeInterface(ref rdaGuid, out var rdaPtr);

            var rda = Marshal.GetObjectForIUnknown(rdaPtr) as IRemoteDebugApplication;

            System.Diagnostics.Debug.WriteLine("process {0} {1} {3}", proc, proc.Id, rda.GetName(out var rdaname), rdaname);

            //languageEngine.SetScriptSite(scriptSite);

            SUCCESS(rda.ConnectDebugger(applicationDebugger));

            

            // --->
            //SUCCESS(rda.GetRootNode(out var dan));
            //SUCCESS(dan.GetName(DOCUMENTNAMETYPE.DOCUMENTNAMETYPE_TITLE, out var name1));

            

            //while (connected)
            //{
            //    System.Threading.Thread.Sleep(10000);
            //}

            //var debugSessionProvider = applicationDebugger as IDebugSessionProvider ?? throw new Exception($"no {nameof(IDebugSessionProvider)}");

            SUCCESS(rda.CauseBreak());

            //SUCCESS(debugSessionProvider.StartDebugSession(debugApplication));

            //scriptSite = new ScriptSite(debugApplication, this) ?? throw new Exception($"no {nameof(IActiveScriptSite)}");


            //dpn2 = provdata.ProgramNodes.Members as IDebugProgramNode2;
        }
        catch { }
    }

    internal void Wait(System.Diagnostics.ProcessStartInfo psi)
    {
        var pdm2_ = Activator.CreateInstance(MSProgramProvider2Type) ?? throw new Exception("no pdm2_");
        var pdm2 = pdm2_ as ActiveDbg.IDebugProgramProvider2My32 ?? throw new Exception("no IDebugProgramProvider2");

        var proc = SuspendedProcess.LaunchProcessSuspended(psi);

        var adprocid = new AD_PROCESS_ID() { dwProcessId = (uint)proc.Process.Id, ProcessIdType = (uint)enum_AD_PROCESS_ID.AD_PROCESS_ID_SYSTEM };

        Callback.Create(out var _callback);

        var flags = enum_PROVIDER_FLAGS.PFLAG_REASON_WATCH | enum_PROVIDER_FLAGS.PFLAG_DEBUGGEE;
        var launchEngine = Guid.Empty;


        // x64 causes Mem Exception
        SUCCESS(pdm2.WatchForProviderEvents(
            flags,
            null,
            adprocid,
            ScriptEngineFilter,
            ref launchEngine,
            _callback
            ));


        proc.ResumeProcess();

        // isn't waiting for sth, just runs through even when notepad.exe called
        _callback.Wait(proc.Process.Id);

        // unreg
        SUCCESS(pdm2.WatchForProviderEvents(
            enum_PROVIDER_FLAGS.PFLAG_NONE,
            null,
            adprocid,
            ScriptEngineFilter,
            ref launchEngine,
            _callback
            ));


    }

    public void CauseBreak()
    {
        debugApplication32?.CauseBreak();
        debugApplication64?.CauseBreak();
    }

    #region " ###################################### Events ###################################### "

    protected override InitializeResponse HandleInitializeRequest(InitializeArguments args)
    {
        DebugWriteMethodeName();
        this.Protocol.SendEvent(new InitializedEvent());

        args.LinesStartAt1 = false;

        return new InitializeResponse()
        {
            SupportsConfigurationDoneRequest = true,
            SupportsTerminateRequest = true,
            SupportSuspendDebuggee = true,
            SupportsDebuggerProperties = true,
            SupportsFunctionBreakpoints = true,
            SupportsInstructionBreakpoints = true,
            SupportsCancelRequest = true,
            SupportsExceptionInfoRequest = true,

        };
    }

    protected override DisconnectResponse HandleDisconnectRequest(DisconnectArguments args)
    {
        DebugWriteMethodeName();
        debugApplication32?.DisconnectDebugger();
        debugApplication64?.DisconnectDebugger();
        SUCCESS(languageEngine.Close());
        return new DisconnectResponse();
    }

    System.Threading.Thread goThread;

    string fileName;
    protected override DAP.Messages.LaunchResponse HandleLaunchRequest(DAP.Messages.LaunchArguments args)
    {
        DebugWriteMethodeName();
        fileName = ((string)args.ConfigurationProperties.GetValueOrDefault("program"));
        var text = System.IO.File.ReadAllText(fileName);
        goThread = new System.Threading.Thread(new ParameterizedThreadStart(Go));
        goThread.Start(text);

        return new LaunchResponse();
    }

    private void Go(object obj)
    {
        DebugWriteMethodeName();
        if (obj is string text)
            this.Parse(text);

        this.Protocol.SendEvent(new ExitedEvent());
        this.Protocol.SendEvent(new TerminatedEvent());
    }

    protected override TerminateResponse HandleTerminateRequest(TerminateArguments args)
    {
        DebugWriteMethodeName();
        System.Environment.Exit(0);
        return new TerminateResponse();
    }

    protected override ConfigurationDoneResponse HandleConfigurationDoneRequest(ConfigurationDoneArguments args)
    {
        DebugWriteMethodeName();
        return new ConfigurationDoneResponse();
    }


    protected override SetBreakpointsResponse HandleSetBreakpointsRequest(SetBreakpointsArguments args)
    {
        DebugWriteMethodeName();
        return new SetBreakpointsResponse(new List<Breakpoint>() { new Breakpoint() });
    }

    protected override void HandleProtocolError(Exception ex)
    {
        DebugWriteMethodeName(ex);
        base.HandleProtocolError(ex);
    }

    protected override SetFunctionBreakpointsResponse HandleSetFunctionBreakpointsRequest(SetFunctionBreakpointsArguments args)
    {
        DebugWriteMethodeName();
        return new SetFunctionBreakpointsResponse();
    }

    protected override SetInstructionBreakpointsResponse HandleSetInstructionBreakpointsRequest(SetInstructionBreakpointsArguments args)
    {
        DebugWriteMethodeName();
        var setInstructionBreakpointsResponse = new SetInstructionBreakpointsResponse();

        return setInstructionBreakpointsResponse;
    }

    protected override ThreadsResponse HandleThreadsRequest(ThreadsArguments args)
    {
        lock (syncobject)
        {
            DebugWriteMethodeName();
            var threadsResponse = new ThreadsResponse();
            threadsResponse.Threads.Add(new DAP.Messages.Thread(1, $"{goThread.ManagedThreadId} {goThread.Name}"));
            // threadsResponse.Threads.Add(new DAP.Messages.Thread(2, "dummy"));
            System.Diagnostics.Debug.WriteLine($"DAP.HandleThreadsRequest {threadsResponse}");
            return threadsResponse;
        }
    }


    protected override StackTraceResponse HandleStackTraceRequest(StackTraceArguments args)
    {
        lock (syncobject)
        {
            DebugWriteMethodeName();
            var str = new StackTraceResponse();
            foreach (var sf in StackFrame.GetFrames(DebugThread))
            {
                str.StackFrames.Add(new DAP.Messages.StackFrame(1, sf.Name, ((int)sf.Line + 1), 1) { Source = new Source() { Path = fileName } });
            }
            str.TotalFrames = str.StackFrames.Count;
            return str;
        }
    }

    Dictionary<int, Variable> dic = new Dictionary<int, Variable>();

    protected override VariablesResponse HandleVariablesRequest(VariablesArguments args)
    {
        lock (syncobject)
        {
            DebugWriteMethodeName(args.VariablesReference);
            var ret = new VariablesResponse();
            int i = dic.Count + 2;
            IEnumerable<Variable> List;
            if (args.VariablesReference > 1)
                List = dic[args.VariablesReference].Members;
            else
                List = Variable.getVariables(this.DebugThread);

            foreach (var v in List)
            {
                if (v.Members.Count() > 0)
                    dic.Add(i, v);

                ret.Variables.Add(new DAP.Messages.Variable(
                    v.Name, v.Value, v.Members.Count() > 0 ? i++ : 0
                ));
            }
            return ret;
        }
    }

    protected override ScopesResponse HandleScopesRequest(ScopesArguments args)
    {
        DebugWriteMethodeName();
        return new ScopesResponse(new List<Scope>() { new Scope("Globals", 1, false) });
    }

    protected override ContinueResponse HandleContinueRequest(ContinueArguments args)
    {
        lock (syncobject)
        {
            DebugWriteMethodeName();
            this.Resume();
            return new ContinueResponse();
        }
    }

    protected override NextResponse HandleNextRequest(NextArguments args)
    {
        lock (syncobject)
        {
            DebugWriteMethodeName();
            this.Resume(BREAKRESUMEACTION.BREAKRESUMEACTION_STEP_OVER);
            return new NextResponse();
        }
    }

    protected override StepInResponse HandleStepInRequest(StepInArguments args)
    {
        lock (syncobject)
        {
            DebugWriteMethodeName();
            this.Resume(BREAKRESUMEACTION.BREAKRESUMEACTION_STEP_INTO);
            return new StepInResponse();
        }

    }

    protected override StepOutResponse HandleStepOutRequest(StepOutArguments args)
    {
        lock (syncobject)
        {
            DebugWriteMethodeName();
            this.Resume(BREAKRESUMEACTION.BREAKRESUMEACTION_STEP_OUT);
            return new StepOutResponse();
        }
    }

    protected override CancelResponse HandleCancelRequest(CancelArguments args)
    {
        lock (syncobject)
        {
            DebugWriteMethodeName();
            this.Resume(BREAKRESUMEACTION.BREAKRESUMEACTION_ABORT);
            return new CancelResponse();
        }
    }

    protected override PauseResponse HandlePauseRequest(PauseArguments args)
    {
        lock (syncobject)
        {
            DebugWriteMethodeName();
            this.debugApplication32?.CauseBreak();
            this.debugApplication64?.CauseBreak();
            return new PauseResponse();
        }
    }

    protected override BreakpointLocationsResponse HandleBreakpointLocationsRequest(BreakpointLocationsArguments args)
    {
        DebugWriteMethodeName();
        return base.HandleBreakpointLocationsRequest(args);
    }

    protected override ExceptionInfoResponse HandleExceptionInfoRequest(ExceptionInfoArguments args)
    {
        DebugWriteMethodeName();
        var exception = this.applicationDebugger.exp[0];
        return new ExceptionInfoResponse()
        {
            Description = $"{exception.bstrDescription}\n{exception.bstrSource} &H{exception.scode.ToString("X")}",
            Code = exception.scode,
        };

    }


    #endregion

}