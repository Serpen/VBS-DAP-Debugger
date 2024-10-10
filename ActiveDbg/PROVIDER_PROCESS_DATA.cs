using System.Runtime.InteropServices;

namespace ActiveDbg
{
    [StructLayout(LayoutKind.Sequential)]
    internal struct PROVIDER_PROCESS_DATA
    {
        public uint Fields;
        public PROGRAM_NODE_ARRAY ProgramNodes;
        public int fIsDebuggerPresent;
    }
}