#region Assembly Microsoft.VisualStudio.Debugger.Interop, Version=8.0.1.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a
// AD7Interop.dll
#endregion

using System.Runtime.InteropServices;

namespace ActiveDbg
{
    [StructLayout(LayoutKind.Sequential)]
    internal struct PROGRAM_NODE_ARRAYMy
    {
        public uint dwCount;

        public /*IDebugProgramNode2*/ IntPtr Members; // Pointer of Pointer **
        
    }
}