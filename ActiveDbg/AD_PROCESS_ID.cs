#region Assembly Microsoft.VisualStudio.Debugger.Interop, Version=8.0.1.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a
// AD7Interop.dll
#endregion


namespace ActiveDbg
{
    public struct AD_PROCESS_IDMy
    {
        public uint ProcessIdType;
        public uint dwProcessId;
        public Guid guidProcessId;
        public uint dwUnused;
    }
}