internal static class Helpers
{
    [System.Diagnostics.DebuggerStepThrough()]
    internal static int SUCCESS(int hresult, string name = "", bool throwException = false)
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
            else
                System.Console.Error.WriteLine($"{name} {hresult} {hresult.ToString("x")} {ex.Message} {file}:{line}");
        }
        return hresult;
    }

    public const int S_OK = 0;
}