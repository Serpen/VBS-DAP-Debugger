namespace ActiveDbg
{
    [Flags]
    public enum DBGPROP_ATTRIB_FLAGS : uint
    {
        //DBGPROP_ATTRIB_NO_ATTRIB = 0x0u,
        //DBGPROP_ATTRIB_VALUE_IS_INVALID = 0x8u,
        //DBGPROP_ATTRIB_VALUE_IS_EXPANDABLE = 0x10u,
        //DBGPROP_ATTRIB_400 = 0x400u,
        //DBGPROP_ATTRIB_VALUE_READONLY = 0x800u,
        //DBGPROP_ATTRIB_ACCESS_PUBLIC = 0x1000u,
        //DBGPROP_ATTRIB_ACCESS_PRIVATE = 0x2000u,
        //DBGPROP_ATTRIB_ACCESS_PROTECTED = 0x4000u,
        //DBGPROP_ATTRIB_ACCESS_FINAL = 0x8000u,
        //DBGPROP_ATTRIB_STORAGE_GLOBAL = 0x10000u,
        //DBGPROP_ATTRIB_STORAGE_STATIC = 0x20000u,
        //DBGPROP_ATTRIB_STORAGE_FIELD = 0x40000u,
        //DBGPROP_ATTRIB_STORAGE_VIRTUAL = 0x80000u,
        //DBGPROP_ATTRIB_TYPE_IS_CONSTANT = 0x100000u,
        //DBGPROP_ATTRIB_TYPE_IS_SYNCHRONIZED = 0x200000u,
        //DBGPROP_ATTRIB_TYPE_IS_VOLATILE = 0x400000u,
        //DBGPROP_ATTRIB_HAS_EXTENDED_ATTRIBS = 0x800000u
        NO_ATTRIB = 0,
        NO_NAME = 0x1,
        NO_TYPE = 0x2,
        NO_VALUE = 0x4,
        VALUE_IS_INVALID = 0x8,
        VALUE_IS_OBJECT = 0x10,
        VALUE_IS_ENUM = 0x20,
        VALUE_IS_CUSTOM = 0x40,
        // OBJECT_IS_EXPANDABLE = 0x70,
        VALUE_HAS_CODE = 0x80,
        TYPE_IS_OBJECT = 0x100,
        // TYPE_IS_EXPANDABLE = 0x100,
        TYPE_HAS_CODE = 0x200,
        SLOT_IS_CATEGORY = 0x400,
        VALUE_READONLY = 0x800,
        ACCESS_PUBLIC = 0x1000,
        ACCESS_PRIVATE = 0x2000,
        ACCESS_PROTECTED = 0x4000,
        ACCESS_FINAL = 0x8000,
        STORAGE_GLOBAL = 0x10000,
        STORAGE_STATIC = 0x20000,
        STORAGE_FIELD = 0x40000,
        STORAGE_VIRTUAL = 0x80000,
        TYPE_IS_CONSTANT = 0x100000,
        TYPE_IS_SYNCHRONIZED = 0x200000,
        TYPE_IS_VOLATILE = 0x400000,
        HAS_EXTENDED_ATTRIBS = 0x800000,
        IS_CLASS = 0x1000000,
        IS_FUNCTION = 0x2000000,
        IS_VARIABLE = 0x4000000,
        IS_PROPERTY = 0x8000000,
        IS_MACRO = 0x10000000,
        IS_TYPE = 0x20000000,
        IS_INHERITED = 0x40000000,
        IS_INTERFACE = 0x80000000
    }

    [Flags]
    public enum DebugPopertyInfoFlags : uint
    {
        //DEBUGPROP_INFO_FULLNAME = 0x1u,
        //DEBUGPROP_INFO_NAME = 0x2u,
        //DEBUGPROP_INFO_TYPE = 0x4u,
        //DEBUGPROP_INFO_VALUE = 0x8u,
        //DEBUGPROP_INFO_ATTRIB = 0x10u,
        //DEBUGPROP_INFO_PROP = 0x20u,
        //DEBUGPROP_INFO_VALUE_AUTOEXPAND = 0x10000u,
        //DEBUGPROP_INFO_NOFUNCEVAL = 0x20000u,
        //DEBUGPROP_INFO_VALUE_RAW = 0x40000u,
        //DEBUGPROP_INFO_VALUE_NO_TOSTRING = 0x80000u,
        //DEBUGPROP_INFO_NO_NONPUBLIC_MEMBERS = 0x100000u,
        //DEBUGPROP_INFO_NONE = 0x0u,
        //DEBUGPROP_INFO_STANDARD = 0x1Eu,
        //DEBUGPROP_INFO_ALL = uint.MaxValue
        PROP_INFO_NAME = 0x1,
        PROP_INFO_TYPE = 0x2,
        PROP_INFO_VALUE = 0x4,
        PROP_INFO_ATTRIBUTES = 0x8,
        PROP_INFO_DEBUGPROP = 0x10,
        PROP_INFO_FULLNAME = 0x20,
        RAW = 0x2000000,
        // PROP_INFO_AUTOEXPAND = 0x8000000,
        // PROP_INFO_STANDARD = (PROP_INFO_NAME | PROP_INFO_TYPE | PROP_INFO_VALUE | PROP_INFO_ATTRIBUTES)
    }


    // https://admhelp.microfocus.com/uft/en/all/VBScript/Content/html/a7c6317d-948f-4bb3-b169-1bbe5b7c7cc5.htm
    [Flags]
    internal enum ScriptItem : uint
    {
        None = 0,
        Unkown1 = 1,
        IsVisible = 2,
        IsSource = 4,
        GlobalMembers = 8,
        Unknown0x10 = 0x10,
        Unknown0x20 = 0x20,
        IsPersistent = 0x40, // 0x00000040
        Unknown0x80 = 0x80,
        Unknown0x100 = 0x100,
        CodeOnly = 0x200, // 0x00000200
        NoCode = 0x400, // 0x00000400
    }

    internal static class EnumPropertyTypes
    {
        internal static readonly Guid IDebugPropertyEnumType_All = new Guid("51973C55-CB0C-11D0-B5C9-00A0244A0E7A");
        internal static readonly Guid IDebugPropertyEnumType_Locals = new Guid("51973C56-CB0C-11D0-B5C9-00A0244A0E7A");
        internal static readonly Guid IDebugPropertyEnumType_Arguments = new Guid("51973C57-CB0C-11D0-B5C9-00A0244A0E7A");
        internal static readonly Guid IDebugPropertyEnumType_LocalsPlusArgs = new Guid("51973C58-CB0C-11D0-B5C9-00A0244A0E7A");
        internal static readonly Guid IDebugPropertyEnumType_Registers = new Guid("51973C59-CB0C-11D0-B5C9-00A0244A0E7A");
    }

    [Flags]
    internal enum ScriptText : uint
    {
        None = 0,
        DelayExecution = 1,
        IsVisible = 2,
        Unknown4 = 4,
        Unknown8 = 8,
        Unknown16 = 0x10,
        IsExpression = 0x20,
        IsPersistent = 0x40,
        HostManageSource = 0x80,
        SCRIPTTEXT_ISXDOMAIN = 0x100,
        SCRIPTTEXT_ISNONUSERCODE = 0x200,
    }

    public enum TEXT_DOC_ATTR : uint
    {
        TEXT_DOC_ATTR_READONLY = 1,
        TEXT_DOC_ATTR_TYPE_PRIMARY = 0x2, // IE?
        TEXT_DOC_ATTR_TYPE_WORKER = 0x4,
        TEXT_DOC_ATTR_TYPE_SCRIPT = 0x8,
    }

    /*
  #define SCRIPTTEXT_DELAYEXECUTION       0x00000001
  #define SCRIPTTEXT_ISVISIBLE            0x00000002
  #define SCRIPTTEXT_ISEXPRESSION         0x00000020
  #define SCRIPTTEXT_ISPERSISTENT         0x00000040
  #define SCRIPTTEXT_HOSTMANAGESSOURCE    0x00000080
  #define SCRIPTTEXT_ISXDOMAIN            0x00000100
  #define SCRIPTTEXT_ISNONUSERCODE        0x00000200
  #define SCRIPTTEXT_ALL_FLAGS            (SCRIPTTEXT_DELAYEXECUTION | \
                                           SCRIPTTEXT_ISVISIBLE | \
                                           SCRIPTTEXT_ISEXPRESSION | \
                                           SCRIPTTEXT_ISPERSISTENT | \
                                           SCRIPTTEXT_HOSTMANAGESSOURCE | \
                                           SCRIPTTEXT_ISXDOMAIN | \
                                           SCRIPTTEXT_ISNONUSERCODE)
    */

    public enum Properties
    {
        SCRIPTPROP_MAJORVERSION  = 1,
        SCRIPTPROP_MINORVERSION  = 2,
        SCRIPTPROP_BUILDNUMBER  = 3,
        SCRIPTPROP_DELAYEDEVENTSINKING  = 0x1000,
        SCRIPTPROP_CATCHEXCEPTION  = 4097,
        SCRIPTPROP_CONVERSIONLCID  = 4098,
        SCRIPTPROP_HOSTSTACKREQUIRED  = 4099,
        SCRIPTPROP_SCRIPTSAREFULLYTRUSTED  = 0x1004, // false
        SCRIPTPROP_INTEGERMODE = 0x3000,
        Unknown61440 = 0xF000,
        SCRIPTPROP_HACK_FIBERSUPPORT  = 0x70000000,
        SCRIPTPROP_HACK_TRIDENTEVENTSINK  = 0x70000001
    }

    public enum SourceTextAttribute : uint
    {
        SPACE = 0,
        KEYWORD = 1,
        COMMENT = 2,
        NONSOURCE = 4,
        OPERATOR = 8,
        NUMBER = 0x10,
        STRING = 0x20,
        FUNCTION_START = 0x40,
        u0x80 = 0x80,
        NAME = 0x100,
        u0x200 = 0x200,
        u0x400 = 0x400,
        u0x4800 = 0x800,
    }

}