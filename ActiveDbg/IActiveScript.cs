using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace ActiveDbg
{
    [ComImport]
    [Guid("BB1A2AE1-A4F9-11CF-8F20-00805F2CD064")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IActiveScript
    {
        [MethodImpl(MethodImplOptions.PreserveSig)]
        int SetScriptSite([In][MarshalAs(UnmanagedType.Interface)] IActiveScriptSite pass);

        [MethodImpl(MethodImplOptions.PreserveSig)]
        int GetScriptSite([In] ref Guid riid, out IntPtr ppvObject);

        [MethodImpl(MethodImplOptions.PreserveSig)]
        int SetScriptState([In]VSDebug.SCRIPTSTATE ss);

        [MethodImpl(MethodImplOptions.PreserveSig)]
        int GetScriptState([Out][MarshalAs(UnmanagedType.LPArray)] VSDebug.SCRIPTSTATE[] pssState);

        [MethodImpl(MethodImplOptions.PreserveSig)]
        int Close();

        [MethodImpl(MethodImplOptions.PreserveSig)]
        int AddNamedItem([In][MarshalAs(UnmanagedType.LPWStr)] string pstrName, [In] uint dwFlags);

        [MethodImpl(MethodImplOptions.PreserveSig)]
        int AddTypeLib([In] ref Guid rguidTypeLib, [In] uint dwMajor, [In] uint dwMinor, [In] uint dwFlags);

        [MethodImpl(MethodImplOptions.PreserveSig)]
        int GetScriptDispatch([In][MarshalAs(UnmanagedType.LPWStr)] string pstrItemName, [MarshalAs(UnmanagedType.IDispatch)] out object ppdisp);

        [MethodImpl(MethodImplOptions.PreserveSig)]
        int GetCurrentScriptThreadID(out uint pstidThread);

        [MethodImpl(MethodImplOptions.PreserveSig)]
        int GetScriptThreadID([In] uint dwWin32ThreadId,  out uint pstidThread);

        [MethodImpl(MethodImplOptions.PreserveSig)]
        int GetScriptThreadState([In] uint stidThread, [Out][MarshalAs(UnmanagedType.LPArray)] VSDebug.SCRIPTTHREADSTATE[] pstsState);

        [MethodImpl(MethodImplOptions.PreserveSig)]
        int InterruptScriptThread([In] uint stidThread, [In][MarshalAs(UnmanagedType.LPArray)] stdole.EXCEPINFO[] pexcepinfo, [In] uint dwFlags);

        [MethodImpl(MethodImplOptions.PreserveSig)]
        int Clone([MarshalAs(UnmanagedType.Interface)] out VSDebug.IActiveScript ppscript);
    }
}