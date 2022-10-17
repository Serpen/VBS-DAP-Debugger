using Microsoft.VisualStudio.Debugger.Interop;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using stdole;

namespace ActiveDbg
{
    [ComImport]
    [Guid("BB1A2AE1-A4F9-11CF-8F20-00805F2CD064")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IActiveScriptMy
    {
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        int SetScriptSite([In][MarshalAs(UnmanagedType.Interface)] IActiveScriptSiteMy pass);

        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        int GetScriptSite([In] ref Guid riid, out IntPtr ppvObject);

        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        int SetScriptState([In][ComAliasName("Microsoft.VisualStudio.Debugger.Interop.SCRIPTSTATE")] SCRIPTSTATE ss);

        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        int GetScriptState([Out][ComAliasName("Microsoft.VisualStudio.Debugger.Interop.SCRIPTSTATE")][MarshalAs(UnmanagedType.LPArray)] SCRIPTSTATE[] pssState);

        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        int Close();

        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        int AddNamedItem([In][MarshalAs(UnmanagedType.LPWStr)] string pstrName, [In] uint dwFlags);

        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        int AddTypeLib([In] ref Guid rguidTypeLib, [In] uint dwMajor, [In] uint dwMinor, [In] uint dwFlags);

        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        int GetScriptDispatch([In][MarshalAs(UnmanagedType.LPWStr)] string pstrItemName, [MarshalAs(UnmanagedType.IDispatch)] out object ppdisp);

        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        int GetCurrentScriptThreadID([ComAliasName("Microsoft.VisualStudio.Debugger.Interop.SCRIPTTHREADID")] out uint pstidThread);

        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        int GetScriptThreadID([In] uint dwWin32ThreadId, [ComAliasName("Microsoft.VisualStudio.Debugger.Interop.SCRIPTTHREADID")] out uint pstidThread);

        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        int GetScriptThreadState([In][ComAliasName("Microsoft.VisualStudio.Debugger.Interop.SCRIPTTHREADID")] uint stidThread, [Out][ComAliasName("Microsoft.VisualStudio.Debugger.Interop.SCRIPTTHREADSTATE")][MarshalAs(UnmanagedType.LPArray)] SCRIPTTHREADSTATE[] pstsState);

        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        int InterruptScriptThread([In][ComAliasName("Microsoft.VisualStudio.Debugger.Interop.SCRIPTTHREADID")] uint stidThread, [In][MarshalAs(UnmanagedType.LPArray)] EXCEPINFO[] pexcepinfo, [In] uint dwFlags);

        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        int Clone([MarshalAs(UnmanagedType.Interface)] out IActiveScript ppscript);
    }
}