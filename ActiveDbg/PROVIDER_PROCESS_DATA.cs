#region Assembly Microsoft.VisualStudio.Debugger.Interop, Version=8.0.1.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a
// AD7Interop.dll
#endregion

using System.Runtime.InteropServices;

namespace ActiveDbg
{
    [StructLayout(LayoutKind.Sequential)]
    internal struct PROVIDER_PROCESS_DATAMy
    {
        public uint Fields;
        public PROGRAM_NODE_ARRAYMy ProgramNodes;
        public int fIsDebuggerPresent;
    }
}