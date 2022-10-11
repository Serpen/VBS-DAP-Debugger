class Program
{
    internal static VbsDebuggerBase vbsbase;

    static System.IO.StreamReader Reader;
    static System.IO.StreamWriter Writer;

    static void Main(string[] args)
    {
        if (args.Length == 0)
            args = new String[] { ".\\Sample.vbs" };

        vbsbase = new VbsDebuggerBase();

        var thread = new Thread(new ParameterizedThreadStart(Go));
        thread.Start(System.IO.File.ReadAllText(args[0]));

        // var pipeServer = new System.IO.Pipes.NamedPipeServerStream("Serpen.vbsdebugger", System.IO.Pipes.PipeDirection.InOut, 1, System.IO.Pipes.PipeTransmissionMode.Message, System.IO.Pipes.PipeOptions.None);
        // Reader = new StreamReader(pipeServer);
        // Writer = new StreamWriter(pipeServer);

        Reader = new StreamReader(System.Console.OpenStandardInput());
        Writer = new StreamWriter(System.Console.OpenStandardOutput()) { AutoFlush = true };

        string choice;
        do
        {
            System.Threading.Thread.Sleep(2000);
            Writer.Write("Action (B/S/V/X/Q): ");
            choice = Reader.ReadLine()?.ToUpper() ?? "";
            if (choice == "X")
            {
                var resThread = new Thread(new ThreadStart(resume));
                resThread.Start();
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
        } while (choice != "" || choice == "Q");
        try
        {
            vbsbase.DebugThread?.Resume(out uint _);
            thread?.Interrupt();
            thread?.Abort();
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
        //vbsbase.Invoke("ScriptEngineMajorVersion & \".\" & ScriptEngineMinorVersion & \".\" & ScriptEngineBuildVersion,");

    }
}