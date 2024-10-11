using static Helpers;
using Marshal = System.Runtime.InteropServices.Marshal;
using System.Diagnostics;

public class VbsDebugAdapter : DAP.DebugAdapterBase, IDisposable
{
    readonly static VSDebug.IProcessDebugManager64 pdm64;
    readonly VSDebug.IDebugApplication64 debugApplication64;

    readonly static VSDebug.IProcessDebugManager32 pdm32;
    readonly VSDebug.IDebugApplication32 debugApplication32;

    internal readonly Debugger applicationDebugger;

    readonly static ActiveDbg.IActiveScript languageEngine;
    readonly ActiveDbg.IActiveScriptSite scriptSite;
    internal VSDebug.IRemoteDebugApplicationThread DebugThread;

    static readonly ActiveDbg.IDebugProgramProvider2 MSProgramProvider2;

    static readonly VSDebug.CONST_GUID_ARRAY ScriptEngineFilter;

    readonly object syncobject = new object();

    bool parsedInited = true;

    static VbsDebugAdapter()
    {
        var PDMtype = Type.GetTypeFromProgID("ProcessDebugManager") ?? throw new Exception("no def ProcessDebugManager");
        var VBScriptType = Type.GetTypeFromProgID("VBScript") ?? throw new Exception("no def VBScriptType");
        var MSProgramProvider2Type = Type.GetTypeFromCLSID(new Guid("{170EC3FC-4E80-40AB-A85A-55900C7C70DE}")) ?? throw new Exception("no def MSProgramProvider2Type"); // ?? throw new Exception("no pdm2type");

        pdm64 = Activator.CreateInstance(PDMtype) as VSDebug.IProcessDebugManager64; // ?? throw new Exception($"no {nameof(IProcessDebugManager64)}");
        pdm32 = Activator.CreateInstance(PDMtype) as VSDebug.IProcessDebugManager32; // ?? throw new Exception($"no {nameof(IProcessDebugManager32)}");
        if (pdm32 is null && pdm64 is null) throw new Exception("no ProcessDebugMananger");

        var VBScript = Activator.CreateInstance(VBScriptType) ?? throw new Exception("no def VBScript");
        languageEngine = VBScript as ActiveDbg.IActiveScript ?? throw new Exception($"no {nameof(ActiveDbg.IActiveScript)}");

        var sefPtr = Marshal.AllocHGlobal((Marshal.SizeOf<Guid>()));

        Guid sefguid = new Guid("{F200A7E7-DEA5-11D0-B854-00A0244A1DE2}");
        Marshal.StructureToPtr(sefguid, sefPtr, false);

        ScriptEngineFilter = new VSDebug.CONST_GUID_ARRAY() { dwCount = 1, Members = sefPtr };

        MSProgramProvider2 = Activator.CreateInstance(MSProgramProvider2Type) as ActiveDbg.IDebugProgramProvider2 ?? throw new Exception($"no {nameof(ActiveDbg.IDebugProgramProvider2)}");
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
            var debugSessionProvider = applicationDebugger as VSDebug.IDebugSessionProvider ?? throw new Exception($"no {nameof(VSDebug.IDebugSessionProvider)}");
            SUCCESS(debugSessionProvider.StartDebugSession(debugApplication64));
            scriptSite = new ScriptSite(debugApplication64, this) ?? throw new Exception($"no {nameof(VSDebug.IActiveScriptSite)}");
        }
        else
        {
            SUCCESS(pdm32.CreateApplication(out debugApplication32));
            ArgumentNullException.ThrowIfNull(debugApplication32);
            SUCCESS(debugApplication32.SetName(nameof(VbsDebugAdapter)));
            SUCCESS(pdm32.AddApplication(debugApplication32, out var cookie));
            applicationDebugger = new Debugger(this);
            var debugSessionProvider = applicationDebugger as VSDebug.IDebugSessionProvider ?? throw new Exception($"no {nameof(VSDebug.IDebugSessionProvider)}");
            SUCCESS(debugSessionProvider.StartDebugSession(debugApplication32));
            scriptSite = new ScriptSite(debugApplication32, this) ?? throw new Exception($"no {nameof(VSDebug.IActiveScriptSite)}");
        }

        System.Diagnostics.Debug.WriteLine("DAP.VbsDebuggerBase InitializeProtocolClient");

        InitializeProtocolClient(System.Console.OpenStandardInput(), System.Console.OpenStandardOutput());

        System.Diagnostics.Debug.WriteLine("DAP.VbsDebuggerBase end");

        // border for -> Parse
    }

    internal void setBreakPoint(uint line)
    {
        var iasd32 = languageEngine as VSDebug.IActiveScriptDebug32;
        var iasd64 = languageEngine as VSDebug.IActiveScriptDebug64;
        VSDebug.IDebugCodeContext dcc;
        VSDebug.IEnumDebugCodeContexts enumDebugCodeContexts;

        if (iasd64 is not null)
            SUCCESS(iasd64.EnumCodeContextsOfPosition(0, line, 0, out enumDebugCodeContexts));
        else
            SUCCESS(iasd32.EnumCodeContextsOfPosition(0, line, 0, out enumDebugCodeContexts));

        SUCCESS(enumDebugCodeContexts.Next(1, out dcc, out var _));

        SUCCESS(dcc.SetBreakPoint(VSDebug.BREAKPOINT_STATE.BREAKPOINT_ENABLED));
    }

    internal void Resume(VSDebug.BREAKRESUMEACTION resumeAction = VSDebug.BREAKRESUMEACTION.BREAKRESUMEACTION_CONTINUE)
    {
        lock (syncobject)
        {
            if (DebugThread is not null)
            {
                DebugThread.GetApplication(out var rdApp);
                rdApp?.ResumeFromBreakPoint(DebugThread, resumeAction, VSDebug.ERRORRESUMEACTION.ERRORRESUMEACTION_SkipErrorStatement);
            }
        }
    }

    public void Parse(string scriptText, IDictionary<string, object> namedItems = null)
    {
        var parser32 = languageEngine as VSDebug.IActiveScriptParse32;
        var parser64 = languageEngine as VSDebug.IActiveScriptParse64;

        if (parser64 is not null)
            Parse64(parser64, scriptText, ActiveDbg.ScriptText.IsVisible, namedItems);
        else
            Parse32(parser32, scriptText, ActiveDbg.ScriptText.IsVisible, namedItems);
    }

    public object Invoke(string scriptText, IDictionary<string, object> namedItems = null)
    {
        var parser32 = languageEngine as VSDebug.IActiveScriptParse32;
        var parser64 = languageEngine as VSDebug.IActiveScriptParse64;

        if (parser64 is not null)
            return Parse64(parser64, scriptText, ActiveDbg.ScriptText.IsVisible | ActiveDbg.ScriptText.IsExpression, namedItems);
        else
            return Parse32(parser32, scriptText, ActiveDbg.ScriptText.IsVisible | ActiveDbg.ScriptText.IsExpression, namedItems);
    }


    object Parse32(VSDebug.IActiveScriptParse32 parser, string scriptText, ActiveDbg.ScriptText flags, IDictionary<string, object> namedItems)
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
            SUCCESS(languageEngine.AddNamedItem(obj.Key, (uint)(ActiveDbg.ScriptItem.IsVisible | ActiveDbg.ScriptItem.IsSource)));
            (scriptSite as ScriptSite).NamedItems.Add(obj.Key, obj.Value);
        }

        SUCCESS(parser.ParseScriptText(scriptText, null, null, null, 0, 0, (uint)flags, out var result, null), throwException: true);
        return result;
    }


    object Parse64(VSDebug.IActiveScriptParse64 parser, string scriptText, ActiveDbg.ScriptText flags, IDictionary<string, object> namedItems)
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
            SUCCESS(languageEngine.AddNamedItem(obj.Key, (uint)(ActiveDbg.ScriptItem.IsVisible | ActiveDbg.ScriptItem.IsSource)));
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

        foreach (var proc in System.Diagnostics.Process.GetProcesses().Where(p => p.ProcessName.Contains("Projekt1vbs", StringComparison.CurrentCultureIgnoreCase)))
        //foreach (var proc in System.Diagnostics.Process.GetProcessesByName("cscript"))
        {
            if (proc.Id == Process.GetCurrentProcess().Id)
                continue;
            System.Diagnostics.Debug.WriteLine("process {0} {1}...", proc, proc.Id);

            var adprocid = new VSDebug.AD_PROCESS_ID() { dwProcessId = (uint)proc.Id, ProcessIdType = (uint)VSDebug.enum_AD_PROCESS_ID.AD_PROCESS_ID_SYSTEM };

            int result = 1;
            try
            {
                result = MSProgramProvider2.GetProviderProcessData(VSDebug.enum_PROVIDER_FLAGS.PFLAG_GET_PROGRAM_NODES, null, adprocid, ScriptEngineFilter, out var provdata);

                SUCCESS(result);

                if (result == 0 && provdata.ProgramNodes.dwCount > 0)
                {
                    result = MSProgramProvider2.GetProviderProcessData(VSDebug.enum_PROVIDER_FLAGS.PFLAG_GET_PROGRAM_NODES | VSDebug.enum_PROVIDER_FLAGS.PFLAG_DEBUGGEE | VSDebug.enum_PROVIDER_FLAGS.PFLAG_ATTACHED_TO_DEBUGGEE , null, adprocid, ScriptEngineFilter, out var provdata2);

                    var dpn2Guid = GetInterfaceGuid(typeof(VSDebug.IDebugProgramNode2));
                    var rdaGuid = GetInterfaceGuid(typeof(VSDebug.IRemoteDebugApplication));

                    var ptrptr = Marshal.ReadIntPtr(provdata2.ProgramNodes.Members);

                    Marshal.QueryInterface(ptrptr, ref dpn2Guid, out var ptr1);

                    var dppn2 = Marshal.GetObjectForIUnknown(ptr1) as VSDebug.IDebugProviderProgramNode2;

                    dppn2.UnmarshalDebuggeeInterface(ref rdaGuid, out var rdaPtr);

                    var rda = Marshal.GetObjectForIUnknown(rdaPtr) as VSDebug.IRemoteDebugApplication;

                    System.Diagnostics.Debug.WriteLine("process {0} {1} {3}", proc, proc.Id, rda.GetName(out var rdaname), rdaname);

                    outlist.Add(proc);
                }
                    
            }
            catch { }
        }
        return outlist;
    }

    internal void Attach(System.Diagnostics.Process proc)
    {
        var adprocid = new VSDebug.AD_PROCESS_ID() { dwProcessId = (uint)proc.Id, ProcessIdType = (uint)VSDebug.enum_AD_PROCESS_ID.AD_PROCESS_ID_SYSTEM };

        int result = 1;
        try
        {
            result = MSProgramProvider2.GetProviderProcessData(
                VSDebug.enum_PROVIDER_FLAGS.PFLAG_GET_PROGRAM_NODES | VSDebug.enum_PROVIDER_FLAGS.PFLAG_DEBUGGEE | VSDebug.enum_PROVIDER_FLAGS.PFLAG_ATTACHED_TO_DEBUGGEE,
                null, adprocid, ScriptEngineFilter, out var provdata);
            

            SUCCESS(result);

            var dpn2Guid = GetInterfaceGuid(typeof(VSDebug.IDebugProgramNode2));
            var rdaGuid = GetInterfaceGuid(typeof(VSDebug.IRemoteDebugApplication));

            var ptrptr = Marshal.ReadIntPtr(provdata.ProgramNodes.Members);

            Marshal.QueryInterface(ptrptr, ref dpn2Guid, out var ptr1);

            var dppn2 = Marshal.GetObjectForIUnknown(ptr1) as VSDebug.IDebugProviderProgramNode2;

            dppn2.UnmarshalDebuggeeInterface(ref rdaGuid, out var rdaPtr);

            var rda = Marshal.GetObjectForIUnknown(rdaPtr) as VSDebug.IRemoteDebugApplication;

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

        var proc = SuspendedProcess.LaunchProcessSuspended(psi);

        var adprocid = new VSDebug.AD_PROCESS_ID() { dwProcessId = (uint)proc.Process.Id, ProcessIdType = (uint)VSDebug.enum_AD_PROCESS_ID.AD_PROCESS_ID_SYSTEM };

        Callback.Create(out var _callback);

        var flags = VSDebug.enum_PROVIDER_FLAGS.PFLAG_REASON_WATCH | VSDebug.enum_PROVIDER_FLAGS.PFLAG_DEBUGGEE;
        var launchEngine = Guid.Empty;


        // x64 causes Mem Exception
        SUCCESS(MSProgramProvider2.WatchForProviderEvents(
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
        SUCCESS(MSProgramProvider2.WatchForProviderEvents(
            VSDebug.enum_PROVIDER_FLAGS.PFLAG_NONE,
            null,
            adprocid,
            ScriptEngineFilter,
            ref launchEngine,
            _callback
            ));


    }

    public void CauseBreak()
    {
        SUCCESS(debugApplication32?.CauseBreak());
        SUCCESS(debugApplication64?.CauseBreak());
    }

    #region " ###################################### Events ###################################### "

    protected override DAP.Messages.InitializeResponse HandleInitializeRequest(DAP.Messages.InitializeArguments args)
    {
        DebugWriteMethodeName();
        this.Protocol.SendEvent(new DAP.Messages.InitializedEvent());

        args.LinesStartAt1 = false;

        return new DAP.Messages.InitializeResponse()
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

    protected override DAP.Messages.DisconnectResponse HandleDisconnectRequest(DAP.Messages.DisconnectArguments args)
    {
        DebugWriteMethodeName();
        debugApplication32?.DisconnectDebugger();
        debugApplication64?.DisconnectDebugger();
        SUCCESS(languageEngine.Close());
        return new DAP.Messages.DisconnectResponse();
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

        return new DAP.Messages.LaunchResponse();
    }

    private void Go(object obj)
    {
        DebugWriteMethodeName();
        if (obj is string text)
            this.Parse(text);

        this.Protocol.SendEvent(new DAP.Messages.ExitedEvent());
        this.Protocol.SendEvent(new DAP.Messages.TerminatedEvent());
    }

    protected override DAP.Messages.TerminateResponse HandleTerminateRequest(DAP.Messages.TerminateArguments args)
    {
        DebugWriteMethodeName();
        System.Environment.Exit(0);
        return new DAP.Messages.TerminateResponse();
    }

    protected override DAP.Messages.ConfigurationDoneResponse HandleConfigurationDoneRequest(DAP.Messages.ConfigurationDoneArguments args)
    {
        DebugWriteMethodeName();
        return new DAP.Messages.ConfigurationDoneResponse();
    }


    protected override DAP.Messages.SetBreakpointsResponse HandleSetBreakpointsRequest(DAP.Messages.SetBreakpointsArguments args)
    {
        DebugWriteMethodeName();
        return new DAP.Messages.SetBreakpointsResponse(new List<DAP.Messages.Breakpoint>() { new DAP.Messages.Breakpoint() });
    }

    protected override void HandleProtocolError(Exception ex)
    {
        DebugWriteMethodeName(ex);
        base.HandleProtocolError(ex);
    }

    protected override DAP.Messages.SetFunctionBreakpointsResponse HandleSetFunctionBreakpointsRequest(DAP.Messages.SetFunctionBreakpointsArguments args)
    {
        DebugWriteMethodeName();
        return new DAP.Messages.SetFunctionBreakpointsResponse();
    }

    protected override DAP.Messages.SetInstructionBreakpointsResponse HandleSetInstructionBreakpointsRequest(DAP.Messages.SetInstructionBreakpointsArguments args)
    {
        DebugWriteMethodeName();
        var setInstructionBreakpointsResponse = new DAP.Messages.SetInstructionBreakpointsResponse();

        return setInstructionBreakpointsResponse;
    }

    protected override DAP.Messages.ThreadsResponse HandleThreadsRequest(DAP.Messages.ThreadsArguments args)
    {
        lock (syncobject)
        {
            DebugWriteMethodeName();
            var threadsResponse = new DAP.Messages.ThreadsResponse();
            threadsResponse.Threads.Add(new DAP.Messages.Thread(1, $"{goThread.ManagedThreadId} {goThread.Name}"));
            // threadsResponse.Threads.Add(new DAP.Messages.Thread(2, "dummy"));
            System.Diagnostics.Debug.WriteLine($"DAP.HandleThreadsRequest {threadsResponse}");
            return threadsResponse;
        }
    }


    protected override DAP.Messages.StackTraceResponse HandleStackTraceRequest(DAP.Messages.StackTraceArguments args)
    {
        lock (syncobject)
        {
            DebugWriteMethodeName();
            var str = new DAP.Messages.StackTraceResponse();
            foreach (var sf in StackFrame.GetFrames(DebugThread))
            {
                str.StackFrames.Add(new DAP.Messages.StackFrame(1, sf.Name, ((int)sf.Line + 1), 1) { Source = new DAP.Messages.Source() { Path = fileName } });
            }
            str.TotalFrames = str.StackFrames.Count;
            return str;
        }
    }

    Dictionary<int, Variable> dic = new Dictionary<int, Variable>();

    protected override DAP.Messages.VariablesResponse HandleVariablesRequest(DAP.Messages.VariablesArguments args)
    {
        lock (syncobject)
        {
            DebugWriteMethodeName(args.VariablesReference);
            var ret = new DAP.Messages.VariablesResponse();
            int i = dic.Count+2;
            IEnumerable<Variable> List;
            if (args.VariablesReference > 1) {
                ret.Variables.Add(new DAP.Messages.Variable("sub", "subval", 0));
                return ret;
                List = dic[args.VariablesReference].Members;
            }
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

    protected override DAP.Messages.ScopesResponse HandleScopesRequest(DAP.Messages.ScopesArguments args)
    {
        lock (syncobject)
        {
            DebugWriteMethodeName();
            return new DAP.Messages.ScopesResponse(new List<DAP.Messages.Scope>() { new DAP.Messages.Scope("Globals", 1, false) });
        }
    }

    protected override DAP.Messages.ContinueResponse HandleContinueRequest(DAP.Messages.ContinueArguments args)
    {
        lock (syncobject)
        {
            DebugWriteMethodeName();
            this.Resume();
            return new DAP.Messages.ContinueResponse();
        }
    }

    protected override DAP.Messages.NextResponse HandleNextRequest(DAP.Messages.NextArguments args)
    {
        lock (syncobject)
        {
            DebugWriteMethodeName();
            this.Resume(VSDebug.BREAKRESUMEACTION.BREAKRESUMEACTION_STEP_OVER);
            return new DAP.Messages.NextResponse();
        }
    }

    protected override DAP.Messages.StepInResponse HandleStepInRequest(DAP.Messages.StepInArguments args)
    {
        lock (syncobject)
        {
            DebugWriteMethodeName();
            this.Resume(VSDebug.BREAKRESUMEACTION.BREAKRESUMEACTION_STEP_INTO);
            return new DAP.Messages.StepInResponse();
        }

    }

    protected override DAP.Messages.StepOutResponse HandleStepOutRequest(DAP.Messages.StepOutArguments args)
    {
        lock (syncobject)
        {
            DebugWriteMethodeName();
            this.Resume(VSDebug.BREAKRESUMEACTION.BREAKRESUMEACTION_STEP_OUT);
            return new DAP.Messages.StepOutResponse();
        }
    }

    protected override DAP.Messages.CancelResponse HandleCancelRequest(DAP.Messages.CancelArguments args)
    {
        lock (syncobject)
        {
            DebugWriteMethodeName();
            this.Resume(VSDebug.BREAKRESUMEACTION.BREAKRESUMEACTION_ABORT);
            return new DAP.Messages.CancelResponse();
        }
    }

    protected override DAP.Messages.PauseResponse HandlePauseRequest(DAP.Messages.PauseArguments args)
    {
        lock (syncobject)
        {
            DebugWriteMethodeName();
            this.debugApplication32?.CauseBreak();
            this.debugApplication64?.CauseBreak();
            return new DAP.Messages.PauseResponse();
        }
    }

    protected override DAP.Messages.BreakpointLocationsResponse HandleBreakpointLocationsRequest(DAP.Messages.BreakpointLocationsArguments args)
    {
        DebugWriteMethodeName();
        return base.HandleBreakpointLocationsRequest(args);
    }

    protected override DAP.Messages.ExceptionInfoResponse HandleExceptionInfoRequest(DAP.Messages.ExceptionInfoArguments args)
    {
        DebugWriteMethodeName();
        var exception = this.applicationDebugger.exp[0];
        return new DAP.Messages.ExceptionInfoResponse()
        {
            Description = $"{exception.bstrDescription}\n{exception.bstrSource} &H{exception.scode.ToString("X")}",
            Code = exception.scode,
        };

    }


    #endregion

}