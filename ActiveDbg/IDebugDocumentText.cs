#region Assembly Microsoft.VisualStudio.Debugger.Interop, Version=8.0.1.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a
// AD7Interop.dll
#endregion

using Microsoft.VisualStudio.Debugger.Interop;

using System.Runtime.InteropServices;

namespace ActiveDbg
{
    [Guid("51973C22-CB0C-11D0-B5C9-00A0244A0E7A")]
    [InterfaceType(1)]
    public interface IDebugDocumentTextMy : IDebugDocument
    {
        int GetName([ComAliasName("Microsoft.VisualStudio.Debugger.Interop.DOCUMENTNAMETYPE")] DOCUMENTNAMETYPE dnt, out string pbstrName);
        int GetDocumentClassId(out Guid pclsidDocument);
        int GetDocumentAttributes([ComAliasName("Microsoft.VisualStudio.Debugger.Interop.TEXT_DOC_ATTR")] out uint ptextdocattr);
        int GetSize(out uint pcNumLines, out uint pcNumChars);
        int GetPositionOfLine(uint cLineNumber, out uint pcCharacterPosition);
        int GetLineOfPosition(uint cCharacterPosition, out uint pcLineNumber, out uint pcCharacterOffsetInLine);
        int GetText(uint cCharacterPosition, IntPtr pcharText, SourceTextAttribute[] pstaTextAttr, out uint pcNumChars, uint cMaxChars);
        int GetPositionOfContext(IDebugDocumentContext psc, out uint pcCharacterPosition, out uint cNumChars);
        int GetContextOfPosition(uint cCharacterPosition, uint cNumChars, out IDebugDocumentContext ppsc);
    }
}