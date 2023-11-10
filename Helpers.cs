internal static class Helpers
{
    [System.Diagnostics.DebuggerStepThrough()]
    internal static int SUCCESS(int hresult, string name = "", bool throwException = false, params int[] ignores)
    {
        if (hresult != 0)
        {
            var st = new System.Diagnostics.StackTrace(true);
            var frame = st.GetFrame(1);
            var file = frame?.GetFileName();
            var line = frame?.GetFileLineNumber();
            var ex = new System.ComponentModel.Win32Exception(hresult);
            if (throwException)
                throw ex;
            else if (!ignores.Contains(hresult))
                System.Console.Error.WriteLine($"{name} {hresult} {hresult.ToString("x")} {ex.Message} {file}:{line}");
        }
        return hresult;
    }

    internal static Guid GetInterfaceGuid(Type Interface)
    => new Guid((Interface.GetCustomAttributes(typeof(System.Runtime.InteropServices.GuidAttribute), false)
            .First() as System.Runtime.InteropServices.GuidAttribute).Value);

    internal static void DebugWriteMethodeName(params object[] args)
    {
        var stack = new System.Diagnostics.StackTrace(true);
        var sframe = stack.GetFrame(1);
        var smethod = sframe.GetMethod();
        System.Diagnostics.Debug.WriteLine(
                   smethod.DeclaringType.Name + "." +
                   smethod.Name +
                   " (" + String.Join(" ", args) + ")");
    }
    public const int S_OK = 0;
}