using System;

namespace ActiveDbg
{
    [Flags]
    public enum enum_DBGPROP_ATTRIB_FLAGS : uint
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
        OBJECT_IS_EXPANDABLE = 0x70,
        VALUE_HAS_CODE = 0x80,
        TYPE_IS_OBJECT = 0x100,
        TYPE_HAS_CODE = 0x200,
        TYPE_IS_EXPANDABLE = 0x100,
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
        public enum enum_DEBUGPROP_INFO_FLAGS : uint
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
        PROP_INFO_FULLNAME = 0x20,
        PROP_INFO_ATTRIBUTES = 0x8,
        PROP_INFO_DEBUGPROP = 0x10,
        PROP_INFO_AUTOEXPAND = 0x8000000,
        PROP_INFO_STANDARD =	( ( ( ( PROP_INFO_NAME | PROP_INFO_TYPE )  | PROP_INFO_VALUE )  | PROP_INFO_ATTRIBUTES )  )
    }
	

// https://admhelp.microfocus.com/uft/en/all/VBScript/Content/html/a7c6317d-948f-4bb3-b169-1bbe5b7c7cc5.htm
  [Flags]
  public enum ScriptItem : uint
  {
    None = 0,
    IsVisible = 2,
    IsSource = 4,
    GlobalMembers = 8,
    IsPersistent = 64, // 0x00000040
    CodeOnly = 512, // 0x00000200
    NoCode = 1024, // 0x00000400
  }

internal static class EnumPropertyTypes
{
    internal static readonly Guid IDebugPropertyEnumType_All =              new Guid("51973C55-CB0C-11D0-B5C9-00A0244A0E7A");
    internal static readonly Guid IDebugPropertyEnumType_Locals =           new Guid("51973C56-CB0C-11D0-B5C9-00A0244A0E7A");
    internal static readonly Guid IDebugPropertyEnumType_Arguments =        new Guid("51973C57-CB0C-11D0-B5C9-00A0244A0E7A");
    internal static readonly Guid IDebugPropertyEnumType_LocalsPlusArgs =   new Guid("51973C58-CB0C-11D0-B5C9-00A0244A0E7A");
    internal static readonly Guid IDebugPropertyEnumType_Registers =        new Guid("51973C59-CB0C-11D0-B5C9-00A0244A0E7A");
}

[Flags]
  public enum ScriptText : uint
  {
    None = 0,
    DelayExecution = 1,
    IsVisible = 2,
    IsExpression = 32, // 0x00000020
    IsPersistent = 64, // 0x00000040
    HostManageSource = 128, // 0x00000080
    SCRIPTTEXT_ISXDOMAIN = 0x00000100,
    SCRIPTTEXT_ISNONUSERCODE = 0x00000200,
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

}