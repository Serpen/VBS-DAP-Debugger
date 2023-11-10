using Microsoft.VisualStudio.Debugger.Interop;
using static Helpers;

public class StackFrame
{
    internal readonly IDebugStackFrame dsf;
    readonly ulong stackStart;

    IDebugDocument debugDocument;

    public static IEnumerable<StackFrame> GetFrames(IRemoteDebugApplicationThread prpt, bool firstOnly = false)
    {
        SUCCESS(prpt.EnumStackFrames(out var enumDebugStackFrames32));

        var enumDebugStackFrames64 = enumDebugStackFrames32 as IEnumDebugStackFrames64;

        if (enumDebugStackFrames64 is not null)
            return GetFrames64(enumDebugStackFrames64, firstOnly);
        else if (enumDebugStackFrames32 is not null)
            return GetFrames32(enumDebugStackFrames32, firstOnly);
        else
            return Array.Empty<StackFrame>();
    }

    static IEnumerable<StackFrame> GetFrames32(IEnumDebugStackFrames enumDebugStackFrames, bool firstOnly = false)
    {
        var retList = new List<StackFrame>();
        var dsfd = new DebugStackFrameDescriptor[1];

        SUCCESS(enumDebugStackFrames.Reset());

        uint fetched = 0;
        do
        {
            SUCCESS(enumDebugStackFrames.Next(1, dsfd, out fetched));

            if (fetched == 1)
                retList.Add(new StackFrame(dsfd[0]));

        } while (fetched > 0 && !firstOnly);
        return retList;
    }

    static IEnumerable<StackFrame> GetFrames64(IEnumDebugStackFrames64 enumDebugStackFrames, bool firstOnly = false)
    {
        var retList = new List<StackFrame>();

        var dsfd = new DebugStackFrameDescriptor64[1];

        SUCCESS(enumDebugStackFrames.Reset());

        uint fetched = 0;
        do
        {
            SUCCESS(enumDebugStackFrames.Next64(1, dsfd, out fetched), ignores: 1);

            if (fetched == 1)
                retList.Add(new StackFrame(dsfd[0]));

        } while (fetched > 0 && !firstOnly);
        return retList;
    }

    public StackFrame(DebugStackFrameDescriptor dsfd)
    {
        dsf = dsfd.pdsf;
        stackStart = dsfd.dwMin;

        SUCCESS(dsf.GetCodeContext(out var debugCodeContext));
        SUCCESS(debugCodeContext.GetDocumentContext(out var debugDocumentContext));
        SUCCESS(debugDocumentContext.GetDocument(out debugDocument));
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
        System.Diagnostics.Debug.WriteLine($"{nameof(StackFrame)}.ctor {this}");
    }
    public StackFrame(DebugStackFrameDescriptor64 dsfd)
    {
        dsf = dsfd.pdsf;
        stackStart = dsfd.dwMin;

        SUCCESS(dsf.GetCodeContext(out var debugCodeContext));
        SUCCESS(debugCodeContext.GetDocumentContext(out var debugDocumentContext));
        SUCCESS(debugDocumentContext.GetDocument(out debugDocument));

        System.Diagnostics.Debug.WriteLine($"{nameof(StackFrame)}.ctor {this}");
    }

    public string Name
    {
        get
        {
            SUCCESS(dsf.GetDescriptionString(0, out string name));
            return name;
        }
    }
    public string NameLong
    {
        get
        {
            SUCCESS(dsf.GetDescriptionString(1, out string name));
            return name;
        }
    }

    public string Language
    {
        get
        {
            SUCCESS(dsf.GetLanguageString(0, out var lang));
            return lang;
        }
    }

    public string NameAppNode
    {
        get
        {
            var debugDocumentText = debugDocument as IDebugDocumentText;

            string name = string.Empty;
            debugDocumentText?.GetName(DOCUMENTNAMETYPE.DOCUMENTNAMETYPE_APPNODE, out name);
            return name;
        }
    }

    public string NameTitle
    {
        get
        {
            var debugDocumentText = debugDocument as IDebugDocumentText;

            string name = string.Empty;
            debugDocumentText?.GetName(DOCUMENTNAMETYPE.DOCUMENTNAMETYPE_TITLE, out name);
            return name;
        }
    }

    public string NameFileTail
    {
        get
        {
            var debugDocumentText = debugDocument as IDebugDocumentText;

            string name = string.Empty;
            debugDocumentText?.GetName(DOCUMENTNAMETYPE.DOCUMENTNAMETYPE_FILE_TAIL, out name);
            return name;
        }
    }

    public Uri NameUrl
    {
        get
        {
            var debugDocumentText = debugDocument as IDebugDocumentText;

            string name = string.Empty;
            debugDocumentText?.GetName(DOCUMENTNAMETYPE.DOCUMENTNAMETYPE_URL, out name);
            if (String.IsNullOrEmpty(name))
                return null;
            else
                return new Uri(name);
        }
    }

    public Guid DocumentClassId
    {
        get
        {
            var debugDocumentText = debugDocument as IDebugDocumentText;

            string name = string.Empty;
            SUCCESS(debugDocumentText.GetDocumentClassId(out var clsid));
            return clsid;
        }
    }

    public ActiveDbg.TEXT_DOC_ATTR DocumentAttributes
    {
        get
        {
            var debugDocumentText = debugDocument as IDebugDocumentText;

            string name = string.Empty;
            SUCCESS(debugDocumentText.GetDocumentAttributes(out var docattr));
            return (ActiveDbg.TEXT_DOC_ATTR)docattr;
        }
    }

    public ulong StackStart => stackStart;

    public override string ToString() => $"&H{StackStart.ToString("X")} {Name} :{Line}";

    public uint Line
    {
        get
        {
            SUCCESS(dsf.GetCodeContext(out var debugCodeContext));
            // SUCCESS(pdsf.GetDebugProperty(out var debugProperty));

            SUCCESS(debugCodeContext.GetDocumentContext(out var debugDocumentContext));

            var debugDocumentText = debugDocument as IDebugDocumentText;

            if (debugDocumentText is null) throw new Exception("");

            // VbsDebuggerBase.pdm.CreateDebugDocumentHelper(0, out var ddh);

            SUCCESS(debugDocumentText.GetPositionOfContext(debugDocumentContext, out var charpos, out var charnum));

            // var strpart = Program.prg.Substring((int)charpos, (int)charnum);

            SUCCESS(debugDocumentText.GetLineOfPosition(charpos, out var line, out var offset));

            return line;
        }
    }
}