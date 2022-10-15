using Microsoft.VisualStudio.Debugger.Interop;
using System.Runtime.InteropServices;

namespace ActiveDbg
{
    [Guid("1959530A-8E53-4E09-AD11-1B7334811CAD")]
    [InterfaceType(1)]
    public interface IDebugProgramProvider2My
    {
        int GetProviderProcessData([ComAliasName("Microsoft.VisualStudio.Debugger.Interop.PROVIDER_FLAGS")] enum_PROVIDER_FLAGS Flags,
            IDebugDefaultPort2 pPort,
            [ComAliasName("Microsoft.VisualStudio.Debugger.Interop.AD_PROCESS_ID")] AD_PROCESS_IDMy ProcessId,
            [ComAliasName("Microsoft.VisualStudio.Debugger.Interop.CONST_GUID_ARRAY")] CONST_GUID_ARRAY EngineFilter,
            [ComAliasName("Microsoft.VisualStudio.Debugger.Interop.PROVIDER_PROCESS_DATA")] out PROVIDER_PROCESS_DATA pProcess);
        int GetProviderProgramNode([ComAliasName("Microsoft.VisualStudio.Debugger.Interop.PROVIDER_FLAGS")] uint Flags, IDebugDefaultPort2 pPort, [ComAliasName("Microsoft.VisualStudio.Debugger.Interop.AD_PROCESS_ID")] AD_PROCESS_IDMy ProcessId, ref Guid guidEngine, [ComAliasName("Microsoft.VisualStudio.Debugger.Interop.UINT64")] ulong programId, out IDebugProgramNode2 ppProgramNode);
        int WatchForProviderEvents([ComAliasName("Microsoft.VisualStudio.Debugger.Interop.PROVIDER_FLAGS")] enum_PROVIDER_FLAGS Flags, IDebugDefaultPort2 pPort, [ComAliasName("Microsoft.VisualStudio.Debugger.Interop.AD_PROCESS_ID")] AD_PROCESS_IDMy ProcessId, [ComAliasName("Microsoft.VisualStudio.Debugger.Interop.CONST_GUID_ARRAY")] CONST_GUID_ARRAY EngineFilter, Guid guidLaunchingEngine, IDebugPortNotify2 pEventCallback);
        int SetLocale(ushort wLangID);
    }
}