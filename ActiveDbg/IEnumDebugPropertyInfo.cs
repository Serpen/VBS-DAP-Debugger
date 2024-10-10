using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace ActiveDbg
{
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("51973C51-CB0C-11D0-B5C9-00A0244A0E7A")] // 2: 6c7072c3-3ac4-408f-a680-fc5a2f96903e
    internal interface IEnumDebugPropertyInfo64
    {
        [MethodImpl(MethodImplOptions.PreserveSig)]
        int Next([In] uint celt, [Out][MarshalAs(UnmanagedType.LPArray)] DebugPropertyInfo64[] pinfo, out uint pcEltsfetched);

        [MethodImpl(MethodImplOptions.PreserveSig)]
        int Skip([In] uint celt);

        [MethodImpl(MethodImplOptions.PreserveSig)]
        int Reset();

        [MethodImpl(MethodImplOptions.PreserveSig)]
        int Clone([MarshalAs(UnmanagedType.Interface)] out IEnumDebugPropertyInfo64 ppepi);

        [MethodImpl(MethodImplOptions.PreserveSig)]
        int GetCount(out uint pcelt);
    }
}