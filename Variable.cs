using ActiveDbg;
using Microsoft.VisualStudio.Debugger.Interop;
using static Helpers;

public class Variable
{

    const uint Radix = 10;
    const DebugPopertyInfoFlags PROP_INFO_STANDARD =
        DebugPopertyInfoFlags.PROP_INFO_NAME |
        DebugPopertyInfoFlags.PROP_INFO_TYPE |
        DebugPopertyInfoFlags.PROP_INFO_VALUE |
        DebugPopertyInfoFlags.PROP_INFO_ATTRIBUTES;

    readonly DebugPropertyInfo64 propertyInfo64;
    readonly DebugPropertyInfo propertyInfo;
    static bool Is64; // readonly -> static

    public static IEnumerable<Variable> getVariables(IRemoteDebugApplicationThread prpt)
    {
        var sf1 = StackFrame.GetFrames(prpt, true).FirstOrDefault();
        if (sf1 is not null)
        {
            SUCCESS(sf1.dsf.GetDebugProperty(out var debugProperty));

            SUCCESS(debugProperty.EnumMembers((uint)(PROP_INFO_STANDARD | DebugPopertyInfoFlags.PROP_INFO_DEBUGPROP), Radix, EnumPropertyTypes.IDebugPropertyEnumType_All, out var enumDebugPropertyInfo32));

            // doesn't works, x64 is still true
            var enumDebugPropertyInfo64 = enumDebugPropertyInfo32 as ActiveDbg.IEnumDebugPropertyInfo64;

            if (Is64 && enumDebugPropertyInfo64 is not null)
                return getVariables64(enumDebugPropertyInfo64);
            else if (enumDebugPropertyInfo32 is not null)
                return getVariables32(enumDebugPropertyInfo32);
        }
        return Array.Empty<Variable>();
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
                if (!isMethodDirty(dpi[i]))
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
                if (!isMethodDirty(dpi[i]))
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

    private static bool isMethodDirty(dynamic dpi) {
        return dpi.m_bstrName == "[Methoden]";
    }

    public string Name => Is64 ? propertyInfo64.m_bstrName : propertyInfo.m_bstrName;
    public string Type => Is64 ? propertyInfo64.m_bstrType : propertyInfo.m_bstrType;
    public string Value => Is64 ? propertyInfo64.m_bstrValue : propertyInfo.m_bstrValue;
    public string Fullname => Is64 ? propertyInfo64.m_bstrFullName : propertyInfo.m_bstrFullName;
    public ActiveDbg.DebugPopertyInfoFlags ValidFields => Is64 ? (ActiveDbg.DebugPopertyInfoFlags)propertyInfo64.m_dwValidFields : (ActiveDbg.DebugPopertyInfoFlags)propertyInfo.m_dwValidFields;
    public ActiveDbg.DBGPROP_ATTRIB_FLAGS Attributes => Is64 ? (ActiveDbg.DBGPROP_ATTRIB_FLAGS)propertyInfo64.m_dwAttrib : (ActiveDbg.DBGPROP_ATTRIB_FLAGS)propertyInfo.m_dwAttrib;

    public IEnumerable<Variable> Members
    {
        get
        {
            if (propertyInfo64.m_pDebugProp is not null)
            {
                SUCCESS(propertyInfo64.m_pDebugProp.EnumMembers((uint)(PROP_INFO_STANDARD | ActiveDbg.DebugPopertyInfoFlags.PROP_INFO_DEBUGPROP), Radix, EnumPropertyTypes.IDebugPropertyEnumType_All, out var enumDebugPropertyInfo32), ignores: -2147467259);

                var enumDebugPropertyInfo64 = enumDebugPropertyInfo32 as ActiveDbg.IEnumDebugPropertyInfo64;

                if (enumDebugPropertyInfo64 is not null)
                    return getVariables64(enumDebugPropertyInfo64);
                else if (enumDebugPropertyInfo32 is not null)
                    return getVariables32(enumDebugPropertyInfo32);

            }
            return Array.Empty<Variable>();
        }
    }
    public override string ToString()
    {
        return $"{Name}{(!String.IsNullOrEmpty(Fullname) ? " [" + Fullname + "]" : "")} As {Type} = {Value} ' {Attributes}";
    }

}