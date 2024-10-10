using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;

namespace ActiveDbg
{
    [ComImport()]
    [Guid("1959530A-8E53-4E09-AD11-1B7334811CAD")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IDebugProgramProvider2
    {
        [MethodImpl(MethodImplOptions.PreserveSig)]
        int GetProviderProcessData(
            VSDebug.enum_PROVIDER_FLAGS Flags,
            VSDebug.IDebugDefaultPort2 pPort,
            VSDebug.AD_PROCESS_ID ProcessId,
            VSDebug.CONST_GUID_ARRAY EngineFilter,
            out PROVIDER_PROCESS_DATA pProcess);

        [MethodImpl(MethodImplOptions.PreserveSig)]
        int GetProviderProgramNode(
            uint Flags,
            VSDebug.IDebugDefaultPort2 pPort, 
            VSDebug.AD_PROCESS_ID ProcessId, 
            ref Guid guidEngine, 
            ulong programId, 
            out VSDebug.IDebugProgramNode2 ppProgramNode);

        [MethodImpl(MethodImplOptions.PreserveSig)]
        int WatchForProviderEvents(
            VSDebug.enum_PROVIDER_FLAGS Flags,
            VSDebug.IDebugDefaultPort2 pPort, 
            VSDebug.AD_PROCESS_ID ProcessId,
            VSDebug.CONST_GUID_ARRAY EngineFilter, 
            ref Guid guidLaunchingEngine,
            VSDebug.IDebugPortNotify2 pEventCallback);
        int SetLocale(ushort wLangID);
    }
}
