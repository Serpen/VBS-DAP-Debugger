using Microsoft.VisualStudio.Debugger.Interop;

public class Callback : IDebugPortNotify2
{
    internal IRemoteDebugApplication rdat;
    
    internal Guid irdapp = new Guid("51973C30-CB0C-11D0-B5C9-00A0244A0E7A");
    public int AddProgramNode(IDebugProgramNode2 pProgramNode)
    {
        throw new NotImplementedException();
        System.Diagnostics.Debug.WriteLine("Script code started inside the IE process.");

        var x = pProgramNode as IDebugProviderProgramNode2;
        x.UnmarshalDebuggeeInterface(ref irdapp, out var intPtr);
        
        rdat = System.Runtime.InteropServices.Marshal.PtrToStructure<IRemoteDebugApplication>(intPtr);

        return 0;
    }

    public int RemoveProgramNode(IDebugProgramNode2 pProgramNode)
    {
        throw new NotImplementedException();
    }

    internal void Wait()
    {
        throw new NotImplementedException();
    }
}