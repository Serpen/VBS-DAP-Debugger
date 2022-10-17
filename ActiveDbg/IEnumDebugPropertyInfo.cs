using Microsoft.VisualStudio.Debugger.Interop;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace ActiveDbg
{
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("51973C51-CB0C-11D0-B5C9-00A0244A0E7A")]
    internal interface IEnumDebugPropertyInfo64
    {
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        int Next([In] uint celt, [Out][MarshalAs(UnmanagedType.LPArray)] DebugPropertyInfo64[] pinfo, out uint pcEltsfetched);

        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        int Skip([In] uint celt);

        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        int Reset();

        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        int Clone([MarshalAs(UnmanagedType.Interface)] out IEnumDebugPropertyInfo ppepi);

        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        int GetCount(out uint pcelt);
    }
}