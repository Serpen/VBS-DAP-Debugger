using DI = Microsoft.VisualStudio.Debugger.Interop;


class Program
{
    internal static VbsDebugAdapter vbsbase;
    static System.IO.StreamReader Reader = new StreamReader(System.Console.OpenStandardInput());
    static System.IO.StreamWriter Writer = new StreamWriter(System.Console.OpenStandardOutput()) { AutoFlush = true };

    public static bool isDAP = false;

    static Thread th = null;

    static void Main(string[] args)
    {
        // throw new ApplicationException("myappx");
        System.Diagnostics.Debug.WriteLine("start program1");
        vbsbase = new VbsDebugAdapter();

        if (args.Length == 1 && args[0] == "-dap")
        {
            Program.isDAP = true;
            vbsbase.Protocol.Run();
        }
        else if (args.Length == 1 && args[0] == "-wait")
        {
            var pi = new System.Diagnostics.ProcessStartInfo("C:\\windows\\system32\\cscript.exe", "//D attach.vbs"); // { WorkingDirectory = "D:\\vs_proj\\VBS-DAP-Debugger\\" };
            vbsbase.Wait(pi);
        }
        else if (args.Length == 1 && args[0] == "-attach")
        {
            foreach (var proc in VbsDebugAdapter.GetScriptProcesses())
            {
                th = new Thread(new ParameterizedThreadStart(Attach));
                th.Start(proc);
                //Attach(proc);
                break;
            }
        }
        else
        {
            if (args.Length == 0)
                args = new String[] { @"D:\Temp\vbscript-files\_allSyntax.vbs" };




            th = new Thread(new ParameterizedThreadStart(Go));
            th.Start(System.IO.File.ReadAllText(args[0]));
        }


        // var pipeServer = new System.IO.Pipes.NamedPipeServerStream("Serpen.vbsdebugger", System.IO.Pipes.PipeDirection.InOut, 1, System.IO.Pipes.PipeTransmissionMode.Message, System.IO.Pipes.PipeOptions.None);
        // Reader = new StreamReader(pipeServer);
        // Writer = new StreamWriter(pipeServer);

        if (true || vbsbase.DebugThread != null)
        {
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
                    {
                        Writer.WriteLine(v);
                        foreach (var mem in v.Members)
                        {
                            Writer.WriteLine(" ." + mem);
                            foreach (var mem2 in mem.Members)
                            {
                                Writer.WriteLine("  ." + mem2);
                            }
                        }
                    }
                }
                else if (choice == "B")
                {
                    vbsbase.CauseBreak();
                }
                else if (choice == "T")
                {

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

    }


    static void resume()
    {
        vbsbase.Resume();
    }

    static void Go(object obj)
    {
        if (obj is string text)
            vbsbase.Parse(text);
        // vbsbase.Invoke("ScriptEngineMajorVersion & \".\" & ScriptEngineMinorVersion & \".\" & ScriptEngineBuildVersion");

    }

    static void Attach(object obj)
    {
        vbsbase.Attach(obj as System.Diagnostics.Process);
    }
}