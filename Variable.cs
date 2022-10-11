using ActiveDbg;
using Microsoft.VisualStudio.Debugger.Interop;
using static Helpers;

public class Variable
{
    private readonly DebugPropertyInfo64 propertyInfo64;
    private readonly DebugPropertyInfo propertyInfo;
    private readonly bool Is64;

    public static IEnumerable<Variable> getVariables(IRemoteDebugApplicationThread prpt )
    {
        
        var retList = new List<Variable>();
        SUCCESS(prpt.EnumStackFrames(out var edsf_native));

#if ARCH64
        var enumDebugStackFrames = edsf_native as IEnumDebugStackFrames64 ?? throw new Exception("no IEnumDebugStackFrames"); ;
#else
        var enumDebugStackFrames = edsf_native;
#endif

        var edsf2 = edsf_native as IEnumDebugStackFrames2;

        SUCCESS(enumDebugStackFrames.Reset());

        uint fetched = 0;

#if ARCH64
        var dstd = new DebugStackFrameDescriptor64[1];
        

        SUCCESS(enumDebugStackFrames.Next64(1, dstd, out fetched));
#else
        var dstd = new DebugStackFrameDescriptor[1];

        SUCCESS(enumDebugStackFrames.Next(1, dstd, out fetched));
#endif

        SUCCESS(dstd[0].pdsf.GetDebugProperty(out var debugProperty));
        
        //var dp2 = debugProperty as IDebugProperty2;
        //var dp3 = debugProperty as IDebugProperty3;

        var flags = (uint)(ActiveDbg.enum_DEBUGPROP_INFO_FLAGS.PROP_INFO_STANDARD);
        //flags = (uint)enum_DBGPROP_INFO_FLAGS.DBGPROP_INFO_TYPE;

        SUCCESS(debugProperty.EnumMembers(flags, 10, EnumPropertyTypes.IDebugPropertyEnumType_All, out var edpi_native));


#if ARCH64
        var enumDebugPropertyInfo = edpi_native as ActiveDbg.IEnumDebugPropertyInfoMy ?? throw new Exception("no IEnumDebugStackFrames"); ;
#else
        var enumDebugPropertyInfo = edpi_native;
#endif

        enumDebugPropertyInfo.Reset();

        enumDebugPropertyInfo.GetCount(out var count);

#if ARCH64
        var dpi = new ActiveDbg.DebugPropertyInfo64[count];
#else
        var dpi = new DebugPropertyInfo[count];
#endif
        do
        {
            enumDebugPropertyInfo.Next(count, dpi, out fetched);

            for (int i = 0; i < fetched; i++)
                retList.Add(new Variable(dpi[i]));
           
        } while (fetched > 0 );
        return retList;
    }

    public Variable(DebugPropertyInfo64 propertyInfo64)
    {
        this.propertyInfo64 = propertyInfo64;
        Is64 = true;
    }

    public Variable(DebugPropertyInfo propertyInfo)
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