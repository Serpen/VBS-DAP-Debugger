using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace ActiveDbg
{
    [ComImport]
    [Guid("DB01A1E3-A42B-11CF-8F20-00805F2CD064")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IActiveScriptSite
    {
        [MethodImpl(MethodImplOptions.PreserveSig)]
        int GetLCID(out uint plcid);

        //IActiveScriptSite
        //[MarshalAs(UnmanagedType.CustomMarshaler, MarshalType = "System.Runtime.InteropServices.CustomMarshalers.TypeToTypeInfoMarshaler, CustomMarshalers, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")]
        //Microsoft.VisualStudio.OLE.Interop.ITypeInfo

        //        Microsoft.VisualStudio.Debugger.Interop.TYPE_INFO
        //https://admhelp.microfocus.com/uft/en/all/VBScript/Content/html/f859ed3b-02c1-4924-99f8-5f5bf1bf8405.htm
        [MethodImpl(MethodImplOptions.PreserveSig)]
        uint GetItemInfo([In][MarshalAs(UnmanagedType.LPWStr)] string pstrName, [In] uint dwReturnMask, [MarshalAs(UnmanagedType.IUnknown)] out object ppiunkItem, /*[MarshalAs(UnmanagedType.CustomMarshaler, MarshalType = "Microsoft.VisualStudio.Debugger.Interop.TYPE_INFO")] */ IntPtr ppti);

        [MethodImpl(MethodImplOptions.PreserveSig)]
        int GetDocVersionString([MarshalAs(UnmanagedType.BStr)] out string pbstrVersion);

        [MethodImpl(MethodImplOptions.PreserveSig)]
        int OnScriptTerminate([In][MarshalAs(UnmanagedType.Struct)] ref object pvarResult, [In][MarshalAs(UnmanagedType.LPArray)] stdole.EXCEPINFO[] pexcepinfo);

        [MethodImpl(MethodImplOptions.PreserveSig)]
        int OnStateChange([In] VSDebug.SCRIPTSTATE ssScriptState);

        [MethodImpl(MethodImplOptions.PreserveSig)]
        int OnScriptError([In][MarshalAs(UnmanagedType.Interface)] VSDebug.IActiveScriptError pscripterror);

        [MethodImpl(MethodImplOptions.PreserveSig)]
        int OnEnterScript();

        [MethodImpl(MethodImplOptions.PreserveSig)]
        int OnLeaveScript();
    }
}