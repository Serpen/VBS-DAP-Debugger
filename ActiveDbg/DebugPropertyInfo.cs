using System.Runtime.InteropServices;

namespace ActiveDbg
{
    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    internal struct DebugPropertyInfo64
    {
        public ulong m_dwValidFields;

        [MarshalAs(UnmanagedType.BStr)]
        public string m_bstrName;

        [MarshalAs(UnmanagedType.BStr)]
        public string m_bstrType;

        [MarshalAs(UnmanagedType.BStr)]
        public string m_bstrValue;

        [MarshalAs(UnmanagedType.BStr)]
        public string m_bstrFullName;

        public ulong m_dwAttrib;

        [MarshalAs(UnmanagedType.Interface)]
        public Microsoft.VisualStudio.Debugger.Interop.IDebugProperty m_pDebugProp;

    }
}