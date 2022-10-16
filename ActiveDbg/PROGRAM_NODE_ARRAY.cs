#region Assembly Microsoft.VisualStudio.Debugger.Interop, Version=8.0.1.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a
// AD7Interop.dll
#endregion

using Microsoft.VisualStudio.Debugger.Interop;
using System.Runtime.InteropServices;

namespace ActiveDbg
{
    [StructLayout(LayoutKind.Sequential)]
    public struct PROGRAM_NODE_ARRAYMy
    {
        public uint dwCount;

        // [ComAliasName("Microsoft.VisualStudio.Debugger.Interop.IDebugProgramNode2")] [MarshalAs(UnmanagedType.IUnknown)]
        public IntPtr Members;
        // public IDebugProgramNode2 Members;
        public uint dwCount2;
        
    }
}