using Microsoft.VisualStudio.Debugger.Interop;
using ActiveDbg;
using static Helpers;
using Marshal = System.Runtime.InteropServices.Marshal;
using DAP = Microsoft.VisualStudio.Shared.VSCodeDebugProtocol;
using Microsoft.VisualStudio.Shared.VSCodeDebugProtocol.Messages;

public class VbsDebuggerBase : DAP.DebugAdapterBase, IDisposable
{

    static internal IProcessDebugManager64 pdm64;
    internal IDebugApplication64 debugApplication64;

    static internal IProcessDebugManager32 pdm32;
    internal IDebugApplication32 debugApplication32;

    private Debugger applicationDebugger;

    static dynamic VBScript;
    static IActiveScriptMy languageEngine;
    private IActiveScriptSiteMy scriptSite;
    internal IRemoteDebugApplicationThread DebugThread;

    static readonly Type MSProgramProvider2Type = Type.GetTypeFromCLSID(new Guid("{170EC3FC-4E80-40AB-A85A-55900C7C70DE}")) ?? throw new Exception("no pdm2type");

    static readonly CONST_GUID_ARRAY ScriptEngineFilter;

    bool parsedInited = true;

    static VbsDebuggerBase()
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
    }
    public VbsDebuggerBase()
    {
        System.Diagnostics.Debug.WriteLine("x_ VbsDebuggerBase");
        //SUCCESS(m_processDebugManager.GetDefaultApplication(out var m_debugApplication));
        if (pdm64 is not null)
        {
            SUCCESS(pdm64.CreateApplication(out debugApplication64));
            ArgumentNullException.ThrowIfNull(debugApplication64);
            SUCCESS(debugApplication64.SetName(nameof(VbsDebuggerBase)));
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
            SUCCESS(debugApplication32.SetName(nameof(VbsDebuggerBase)));
            SUCCESS(pdm32.AddApplication(debugApplication32, out var cookie));
            applicationDebugger = new Debugger(this);
            var debugSessionProvider = applicationDebugger as IDebugSessionProvider ?? throw new Exception($"no {nameof(IDebugSessionProvider)}");
            SUCCESS(debugSessionProvider.StartDebugSession(debugApplication32));
            scriptSite = new ScriptSite(debugApplication32, this) ?? throw new Exception($"no {nameof(IActiveScriptSite)}");
        }

        System.Threading.Thread.Sleep(1000);

        System.Diagnostics.Debug.WriteLine("x_ VbsDebuggerBase InitializeProtocolClient");

        InitializeProtocolClient(System.Console.OpenStandardInput(), System.Console.OpenStandardOutput());

        System.Diagnostics.Debug.WriteLine("x_ VbsDebuggerBase end");

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

        SUCCESS(parser.ParseScriptText(scriptText, null, null, null, 0, 0, (uint)flags, out var result, null), throwException: true);
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

    #region Events
    protected override InitializeResponse HandleInitializeRequest(InitializeArguments arguments)
    {
        System.Diagnostics.Debug.WriteLine("x_ InitializeResponse");
        this.Protocol.SendEvent(new InitializedEvent());

        System.Threading.Thread.Sleep(1000);

        arguments.LinesStartAt1 = false;

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

    protected override DisconnectResponse HandleDisconnectRequest(DisconnectArguments arguments)
    {
        System.Diagnostics.Debug.WriteLine("x_ HandleDisconnectRequest");
        debugApplication32?.DisconnectDebugger();
        debugApplication64?.DisconnectDebugger();
        SUCCESS(languageEngine.Close());
        return base.HandleDisconnectRequest(arguments);
    }

    System.Threading.Thread goThread;

    protected override DAP.Messages.LaunchResponse HandleLaunchRequest(DAP.Messages.LaunchArguments args)
    {
        System.Diagnostics.Debug.WriteLine("x_ HandleLaunchRequest");
        var fileName = ((string)args.ConfigurationProperties.GetValueOrDefault("program"));
        var text = System.IO.File.ReadAllText(fileName);
        goThread = new System.Threading.Thread(new ParameterizedThreadStart(Go));
        System.Diagnostics.Debug.WriteLine("x_ HandleLaunchRequest 1");
        goThread.Start(text);
        System.Diagnostics.Debug.WriteLine("x_ HandleLaunchRequest 2");

        return new LaunchResponse();
    }

    private void Go(object? obj)
    {
        System.Diagnostics.Debug.WriteLine("x_ Go");
        if (obj is string text)
            this.Parse(text);

        this.Protocol.SendEvent(new ExitedEvent());
        this.Protocol.SendEvent(new TerminatedEvent());
    }

    protected override TerminateResponse HandleTerminateRequest(TerminateArguments arguments)
    {
        System.Diagnostics.Debug.WriteLine("x_ HandleTerminateRequest");
        System.Environment.Exit(0);
        return new TerminateResponse();
    }

    protected override ConfigurationDoneResponse HandleConfigurationDoneRequest(ConfigurationDoneArguments arguments)
    {
        System.Diagnostics.Debug.WriteLine("x_ HandleConfigurationDoneRequest");
        return new ConfigurationDoneResponse();
    }


    protected override SetBreakpointsResponse HandleSetBreakpointsRequest(SetBreakpointsArguments arguments)
    {
        System.Diagnostics.Debug.WriteLine("x_ SetBreakpointsResponse");
        return new SetBreakpointsResponse(new List<Breakpoint>() { new Breakpoint() });
    }

    protected override void HandleProtocolError(Exception ex)
    {
        System.Diagnostics.Debug.WriteLine($"x_ HandleProtocolError {ex}");
        base.HandleProtocolError(ex);
    }

    protected override SetFunctionBreakpointsResponse HandleSetFunctionBreakpointsRequest(SetFunctionBreakpointsArguments arguments)
    {
        System.Diagnostics.Debug.WriteLine("x_ HandleSetFunctionBreakpointsRequest");
        return new SetFunctionBreakpointsResponse();
        return base.HandleSetFunctionBreakpointsRequest(arguments);
    }

    protected override SetInstructionBreakpointsResponse HandleSetInstructionBreakpointsRequest(SetInstructionBreakpointsArguments arguments)
    {
        System.Diagnostics.Debug.WriteLine("x_ HandleSetInstructionBreakpointsRequest");
        var setInstructionBreakpointsResponse = new SetInstructionBreakpointsResponse();

        return setInstructionBreakpointsResponse;
        return base.HandleSetInstructionBreakpointsRequest(arguments);
    }

    protected override ThreadsResponse HandleThreadsRequest(ThreadsArguments arguments)
    {
        System.Diagnostics.Debug.WriteLine($"x_ HandleThreadsRequest {arguments}");
        var threadsResponse = new ThreadsResponse();
        threadsResponse.Threads.Add(new DAP.Messages.Thread(1, $"{goThread.ManagedThreadId} {goThread.Name}"));
        // threadsResponse.Threads.Add(new DAP.Messages.Thread(2, "dummy"));
        System.Diagnostics.Debug.WriteLine($"x_ HandleThreadsRequest {threadsResponse}");
        return threadsResponse;
    }


    protected override StackTraceResponse HandleStackTraceRequest(StackTraceArguments arguments)
    {
        System.Diagnostics.Debug.WriteLine("x_ HandleStackTraceRequest");
        var str = new StackTraceResponse();
        foreach (var sf in StackFrame.GetFrames(DebugThread))
        {
            str.StackFrames.Add(new DAP.Messages.StackFrame(1, sf.Name, ((int)sf.Line + 1), 1));
        }
        str.TotalFrames = str.StackFrames.Count;
        return str;
    }

    protected override VariablesResponse HandleVariablesRequest(VariablesArguments args)
    {
        System.Diagnostics.Debug.WriteLine($"x_ VariablesArguments {args.VariablesReference}");
        var ret = new VariablesResponse();
        int i = 2;
        if (args.VariablesReference > 1) return ret;
        foreach (var v in Variable.getVariables(this.DebugThread))
        {
            ret.Variables.Add(new DAP.Messages.Variable(
                v.Name, v.Value, v.Members.Count() > 0 ? i++ : 0
            ));
            if (i == args.VariablesReference)
            {
                ret.Variables.Clear();
                foreach (var v2 in v.Members)
                {
                    ret.Variables.Add(new DAP.Messages.Variable(v2.Name, v2.Value, 0));
                }
                return ret;
            }
        }
        return ret;
    }

    protected override ScopesResponse HandleScopesRequest(ScopesArguments arguments)
    {
        System.Diagnostics.Debug.WriteLine($"x_ HandleScopesRequest");

        return new ScopesResponse(new List<Scope>() { new Scope("Globals", 1, false) });
    }

    protected override ContinueResponse HandleContinueRequest(ContinueArguments arguments)
    {
        System.Diagnostics.Debug.WriteLine($"x_ HandleContinueRequest");
        this.Resume();
        return new ContinueResponse();
    }

    protected override NextResponse HandleNextRequest(NextArguments arguments)
    {
        System.Diagnostics.Debug.WriteLine($"x_ HandleStepInRequest");
        this.Resume(BREAKRESUMEACTION.BREAKRESUMEACTION_STEP_OVER);
        return new NextResponse();
    }

    protected override StepInResponse HandleStepInRequest(StepInArguments arguments)
    {
        System.Diagnostics.Debug.WriteLine($"x_ HandleStepInRequest");
        this.Resume(BREAKRESUMEACTION.BREAKRESUMEACTION_STEP_INTO);
        return new StepInResponse();
    }

    protected override StepOutResponse HandleStepOutRequest(StepOutArguments arguments)
    {
        System.Diagnostics.Debug.WriteLine($"x_ HandleStepOutRequest");
        this.Resume(BREAKRESUMEACTION.BREAKRESUMEACTION_STEP_OUT);
        return new StepOutResponse();
    }

    protected override CancelResponse HandleCancelRequest(CancelArguments arguments)
    {
        System.Diagnostics.Debug.WriteLine($"x_ HandleCancelRequest");
        this.Resume(BREAKRESUMEACTION.BREAKRESUMEACTION_ABORT);
        return new CancelResponse();
    }

    protected override PauseResponse HandlePauseRequest(PauseArguments arguments)
    {
        System.Diagnostics.Debug.WriteLine($"x_ HandlePauseRequest");
        this.debugApplication32?.CauseBreak();
        this.debugApplication64?.CauseBreak();
        return new PauseResponse();
    }

    protected override BreakpointLocationsResponse HandleBreakpointLocationsRequest(BreakpointLocationsArguments arguments)
    {
        System.Diagnostics.Debug.WriteLine($"x_ HandleBreakpointLocationsRequest");
        return base.HandleBreakpointLocationsRequest(arguments);
    }

    protected override ExceptionInfoResponse HandleExceptionInfoRequest(ExceptionInfoArguments arguments)
    {
        System.Diagnostics.Debug.WriteLine($"x_ HandleExceptionInfoRequest");
        return new ExceptionInfoResponse()
        {
            Description = this.applicationDebugger.exp[0].bstrDescription,
            Code = this.applicationDebugger.exp[0].wCode,
        };
        
    }


    #endregion

}