using System.Runtime.InteropServices;

namespace ActiveDbg
{
    [StructLayout(LayoutKind.Sequential)]
    internal struct PROGRAM_NODE_ARRAY
    {
        public uint dwCount;

        public /*IDebugProgramNode2*/ IntPtr Members; // Pointer of Pointer **

    }
}