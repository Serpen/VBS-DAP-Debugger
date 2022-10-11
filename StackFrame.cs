using Microsoft.VisualStudio.Debugger.Interop;
using static Helpers;


public class StackFrame
{
    public static IEnumerable<StackFrame> GetFrames(IRemoteDebugApplicationThread prpt)
    {
        var retList = new List<StackFrame>();
        SUCCESS(prpt.EnumStackFrames(out var edsf_native));

#if ARCH64
        var enumDebugStackFrames = edsf_native as IEnumDebugStackFrames64 ?? throw new Exception("no IEnumDebugStackFrames"); ;
#else
        var enumDebugStackFrames = edsf_native;
#endif

        SUCCESS(enumDebugStackFrames.Reset());

#if ARCH64
        var dsfd = new DebugStackFrameDescriptor64[5];
#else
        var dsfd = new DebugStackFrameDescriptor[5];
#endif
        uint fetched = 0;
        do
        {
#if ARCH64
            SUCCESS(enumDebugStackFrames.Next64((uint)dsfd.Length, dsfd, out fetched));
#else
            SUCCESS(enumDebugStackFrames.Next((uint)dsfd.Length, dsfd, out fetched));
#endif
            for (int i = 0; i < fetched; i++)
                retList.Add(new StackFrame(dsfd[i]));


        } while (fetched > 0 && fetched == dsfd.Length);
        return retList;
    }


    private readonly IDebugStackFrame pdsf;
    private readonly IDebugApplicationThread thread;
    private readonly ulong stackstart;

    public StackFrame(DebugStackFrameDescriptor dsfd)
    {
        this.pdsf = dsfd.pdsf;

        /*
         * SUCCESS(pdsf.GetThread(out this.thread));
         * -2147221163 80040155 Schnittstelle nicht registriert
         * Ausnahme ausgel�st bei 0x7769E5A2 (KernelBase.dll) in vbs-atl-cs.exe: WinRT originate error - 0x80040155 : 'Die Proxyregistrierung f�r die IID {51973C38-CB0C-11D0-B5C9-00A0244A0E7A} wurde nicht gefunden.'.
            marshal.cxx(1284)\combase.dll!773F51EF: (caller: 773F4D5C) ReturnHr(1) tid(5b60) 80040155 Schnittstelle nicht registriert
            Msg:[Failed to marshal with IID={51973C38-CB0C-11D0-B5C9-00A0244A0E7A}] 
            onecore\com\combase\dcomrem\marshal.cxx(1179)\combase.dll!773F4EF5: (caller: 773E1C94) LogHr(1) tid(5b60) 80040155 Schnittstelle nicht registriert
            onecore\com\combase\dcomrem\marshal.cxx(1119)\combase.dll!773E1CEA: (caller: 773DC534) ReturnHr(2) tid(5b60) 80040155 Schnittstelle nicht registriert
            Ausnahme ausgel�st bei 0x7769E5A2 (KernelBase.dll) in vbs-atl-cs.exe: 0x80040155: Schnittstelle nicht registriert.
            Ausnahme ausgel�st bei 0x7769E5A2 (KernelBase.dll) in vbs-atl-cs.exe: 0x80040155: Schnittstelle nicht registriert.
            pdsf.GetThread(out this.thread);
         */

        stackstart = dsfd.dwMin;
    }
    public StackFrame(DebugStackFrameDescriptor64 dsfd)
    {
        this.pdsf = dsfd.pdsf;
        /*
         * SUCCESS(pdsf.GetThread(out this.thread));
         * -2147221163 80040155 Schnittstelle nicht registriert
         * Ausnahme ausgel�st bei 0x7769E5A2 (KernelBase.dll) in vbs-atl-cs.exe: WinRT originate error - 0x80040155 : 'Die Proxyregistrierung f�r die IID {51973C38-CB0C-11D0-B5C9-00A0244A0E7A} wurde nicht gefunden.'.
            marshal.cxx(1284)\combase.dll!773F51EF: (caller: 773F4D5C) ReturnHr(1) tid(5b60) 80040155 Schnittstelle nicht registriert
            Msg:[Failed to marshal with IID={51973C38-CB0C-11D0-B5C9-00A0244A0E7A}] 
            onecore\com\combase\dcomrem\marshal.cxx(1179)\combase.dll!773F4EF5: (caller: 773E1C94) LogHr(1) tid(5b60) 80040155 Schnittstelle nicht registriert
            onecore\com\combase\dcomrem\marshal.cxx(1119)\combase.dll!773E1CEA: (caller: 773DC534) ReturnHr(2) tid(5b60) 80040155 Schnittstelle nicht registriert
            Ausnahme ausgel�st bei 0x7769E5A2 (KernelBase.dll) in vbs-atl-cs.exe: 0x80040155: Schnittstelle nicht registriert.
            Ausnahme ausgel�st bei 0x7769E5A2 (KernelBase.dll) in vbs-atl-cs.exe: 0x80040155: Schnittstelle nicht registriert.
            pdsf.GetThread(out this.thread);
         */
        stackstart = dsfd.dwMin;
    }

    public string Name
    {
        get
        {
            pdsf.GetDescriptionString(0, out string name);
            return name;
        }
    }

    public ulong StackStart { get => stackstart; }

    public uint ThreadID
    {
        get
        {
            uint threadid = 0;
            thread?.GetSystemThreadId(out threadid);
            return threadid;
        }
    }


    public System.Threading.ThreadState ThreadState
    {
        get
        {
            uint state = 8;
            thread?.GetState(out state);
            return (System.Threading.ThreadState)state;
        }
    }

    public override string ToString()
    {
        return $"&H{StackStart.ToString("x")} {Name} :{Line}";
    }

    public string Line
    {
        get
        {
            SUCCESS(pdsf.GetCodeContext(out var debugCodeContext));
            //SUCCESS(pdsf.GetDebugProperty(out var debugProperty));

            //SUCCESS(pdsf.GetDescriptionString(0, out var descstring0));
            //SUCCESS(pdsf.GetDescriptionString(1, out var descstring1));
            //SUCCESS(pdsf.GetLanguageString(0, out var lang));

            SUCCESS(debugCodeContext.GetDocumentContext(out var debugDocumentContext));

            SUCCESS(debugDocumentContext.GetDocument(out var debugDocument));

            var debugDocumentText = debugDocument as IDebugDocumentText;

            if (debugDocumentText is null) throw new Exception("");

            //SUCCESS(debugDocumentText.GetName(DOCUMENTNAMETYPE.DOCUMENTNAMETYPE_URL, out var name3));
            //SUCCESS(debugDocumentText.GetName(DOCUMENTNAMETYPE.DOCUMENTNAMETYPE_APPNODE, out var name0));
            //SUCCESS(debugDocumentText.GetName(DOCUMENTNAMETYPE.DOCUMENTNAMETYPE_TITLE, out var name1));
            //SUCCESS(debugDocumentText.GetName(DOCUMENTNAMETYPE.DOCUMENTNAMETYPE_FILE_TAIL, out var name4));

            //SUCCESS(debugDocumentText.GetDocumentAttributes(out var docattr));
            //SUCCESS(debugDocumentText.GetDocumentClassId(out var clsid));

            VbsDebuggerBase.pdm.CreateDebugDocumentHelper(0, out var ddh);
            VbsDebuggerBase.pdm.CreateApplication(out var x);

            SUCCESS(debugDocumentText.GetPositionOfContext(debugDocumentContext, out var charpos, out var charnum));

            //var strpart = Program.prg.Substring((int)charpos, (int)charnum);

            SUCCESS(debugDocumentText.GetLineOfPosition(charpos, out var line, out var offset));

            return line.ToString();
        }
    }
}