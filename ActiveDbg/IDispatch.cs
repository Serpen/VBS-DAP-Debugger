using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

[ComImport]
[Guid("00020400-0000-0000-c000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IDispatch
{
    [PreserveSig]
    int GetTypeInfoCount(
        [Out] out uint count
    );

    [PreserveSig]
    int GetTypeInfo(
        [In] uint index,
        [In] int lcid,
        [Out][MarshalAs(UnmanagedType.Interface)] out ITypeInfo typeInfo
    );

    [PreserveSig]
    int GetIDsOfNames(
        [In] ref Guid iid,
        [In][MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr, SizeParamIndex = 2)] string[] names,
        [In] uint count,
        [In] int lcid,
        [Out][MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 2)] int[] dispids
    );

    [PreserveSig]
    int Invoke(
        [In] int dispid,
        [In] ref Guid iid,
        [In] int lcid,
        [In] int flags,
        [In] ref DISPPARAMS args,
        [In] IntPtr pVarResult,
        [Out] out EXCEPINFO excepInfo,
        [Out] out uint argErr
    );
}