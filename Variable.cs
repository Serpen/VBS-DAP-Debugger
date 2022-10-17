using ActiveDbg;
using Microsoft.VisualStudio.Debugger.Interop;
using static Helpers;

public class Variable
{
    private readonly DebugPropertyInfo64 propertyInfo64;
    private readonly DebugPropertyInfo propertyInfo;
    private readonly bool Is64;

    public static IEnumerable<Variable> getVariables(IRemoteDebugApplicationThread prpt)
    {
        var sf1 = StackFrame.GetFrames(prpt, true).First();

        SUCCESS(sf1.dsf.GetDebugProperty(out var debugProperty));

        SUCCESS(debugProperty.EnumMembers((uint)(ActiveDbg.enum_DEBUGPROP_INFO_FLAGS.PROP_INFO_STANDARD), 10, EnumPropertyTypes.IDebugPropertyEnumType_All, out var enumDebugPropertyInfo32));

        var enumDebugPropertyInfo64 = enumDebugPropertyInfo32 as ActiveDbg.IEnumDebugPropertyInfo64;

        if (enumDebugPropertyInfo64 is not null)
            return getVariables64(enumDebugPropertyInfo64);
        else
            return getVariables32(enumDebugPropertyInfo32);
    }

    static IEnumerable<Variable> getVariables32(IEnumDebugPropertyInfo edpi32)
    {
        var retList = new List<Variable>();

        edpi32.Reset();

        edpi32.GetCount(out var count);

        var dpi = new DebugPropertyInfo[count];

        uint fetched = 0;
        do
        {
            edpi32.Next(count, dpi, out fetched);

            for (int i = 0; i < fetched; i++)
                retList.Add(new Variable(dpi[i]));

        } while (fetched > 0);
        return retList;
    }

    static IEnumerable<Variable> getVariables64(ActiveDbg.IEnumDebugPropertyInfo64 edpi64)
    {
        var retList = new List<Variable>();

        edpi64.Reset();

        edpi64.GetCount(out var count);

        var dpi = new ActiveDbg.DebugPropertyInfo64[count];

        uint fetched = 0;
        do
        {
            edpi64.Next(count, dpi, out fetched);

            for (int i = 0; i < fetched; i++)
                retList.Add(new Variable(dpi[i]));

        } while (fetched > 0);
        return retList;
    }

    internal Variable(DebugPropertyInfo64 propertyInfo64)
    {
        this.propertyInfo64 = propertyInfo64;
        Is64 = true;
    }

    internal Variable(DebugPropertyInfo propertyInfo)
    {
        this.propertyInfo = propertyInfo;
        Is64 = false;
    }

    public string Name => Is64 ? propertyInfo64.m_bstrName : propertyInfo.m_bstrName;
    public string Type => Is64 ? propertyInfo64.m_bstrType : propertyInfo.m_bstrType;
    public string Value => Is64 ? propertyInfo64.m_bstrValue : propertyInfo.m_bstrValue;
    public string Fullname => Is64 ? propertyInfo64.m_bstrFullName : propertyInfo.m_bstrFullName;
    public ActiveDbg.enum_DEBUGPROP_INFO_FLAGS ValidFields => Is64 ? (ActiveDbg.enum_DEBUGPROP_INFO_FLAGS)propertyInfo64.m_dwValidFields : (ActiveDbg.enum_DEBUGPROP_INFO_FLAGS)propertyInfo.m_dwValidFields;
    public ActiveDbg.enum_DBGPROP_ATTRIB_FLAGS Attributes => Is64 ? (ActiveDbg.enum_DBGPROP_ATTRIB_FLAGS)propertyInfo64.m_dwAttrib : (ActiveDbg.enum_DBGPROP_ATTRIB_FLAGS)propertyInfo.m_dwAttrib;

    public override string ToString()
    {
        return $"{Name}{(!String.IsNullOrEmpty(Fullname) ? " [" + Fullname + "]" : "")} As {Type} = {Value} ' {Attributes}";
    }

}