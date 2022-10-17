class Program
{
    internal static VbsDebuggerBase vbsbase;

    static System.IO.StreamReader Reader;
    static System.IO.StreamWriter Writer;

    static void Main(string[] args)
    {
        System.Console.WriteLine("x");
        if (args.Length == 0)
            args = new String[] { ".\\Sample.vbs" };

        vbsbase = new VbsDebuggerBase();

        Thread? th = null;

        if (false)
        {
            foreach (var proc in VbsDebuggerBase.GetScriptProcesses())
            {
                // th = new Thread(new ParameterizedThreadStart(Attach));
                // th.Start(proc);
                Attach(proc);
                break;
            }
        }
        else
        {
            th = new Thread(new ParameterizedThreadStart(Go));
            th.Start(System.IO.File.ReadAllText(args[0]));
        }

        // var pipeServer = new System.IO.Pipes.NamedPipeServerStream("Serpen.vbsdebugger", System.IO.Pipes.PipeDirection.InOut, 1, System.IO.Pipes.PipeTransmissionMode.Message, System.IO.Pipes.PipeOptions.None);
        // Reader = new StreamReader(pipeServer);
        // Writer = new StreamWriter(pipeServer);

        Reader = new StreamReader(System.Console.OpenStandardInput());
        Writer = new StreamWriter(System.Console.OpenStandardOutput()) { AutoFlush = true };

        string choice;
        do
        {
            System.Threading.Thread.Sleep(2000);
            Writer.Write("Action (B/S/V/R/Q): ");
            choice = Reader.ReadLine()?.ToUpper() ?? "";
            if (choice == "R")
            {
                // var resThread = new Thread(new ThreadStart(resume));
                // resThread.Start();
                resume();
            }
            else if (choice == "F" | choice == "S")
            {
                foreach (var sf in StackFrame.GetFrames(vbsbase.DebugThread))
                    Writer.WriteLine(sf);
            }
            else if (choice == "V")
            {
                foreach (var v in Variable.getVariables(vbsbase.DebugThread))
                    Writer.WriteLine(v);
            }
            else if (choice == "B")
            {
                System.Console.Write("Breakpoint-Line: ");
                if (System.UInt32.TryParse(Reader.ReadLine(), out var line))
                    vbsbase.setBreakPoint(line);
            }
        } while (choice != "" && choice != "Q");
        try
        {
            vbsbase.DebugThread?.Resume(out uint _);
            th?.Interrupt();
            th?.Abort();
        }
        catch
        {

        }
    }

    static void resume()
    {
        vbsbase.Resume();
    }

    static void Go(object? obj)
    {
        if (obj is string text)
            vbsbase.Parse(text);
        // vbsbase.Invoke("ScriptEngineMajorVersion & \".\" & ScriptEngineMinorVersion & \".\" & ScriptEngineBuildVersion");

    }

    static void Attach(object? obj)
    {
        vbsbase.Attach(obj as System.Diagnostics.Process);
    }
}