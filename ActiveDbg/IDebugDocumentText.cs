using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace ActiveDbg
{
    [ComImport()]
    [Guid("51973C22-CB0C-11D0-B5C9-00A0244A0E7A")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDebugDocumentText : IDebugDocument
    {
        int GetName( DOCUMENTNAMETYPE dnt, out string pbstrName);
        int GetDocumentClassId(out Guid pclsidDocument);
        int GetDocumentAttributes(out uint ptextdocattr);
        int GetSize(out uint pcNumLines, out uint pcNumChars);
        int GetPositionOfLine(uint cLineNumber, out uint pcCharacterPosition);
        int GetLineOfPosition(uint cCharacterPosition, out uint pcLineNumber, out uint pcCharacterOffsetInLine);
        int GetText(uint cCharacterPosition, IntPtr pcharText, SourceTextAttribute[] pstaTextAttr, out uint pcNumChars, uint cMaxChars);
        int GetPositionOfContext(IDebugDocumentContext psc, out uint pcCharacterPosition, out uint cNumChars);
        int GetContextOfPosition(uint cCharacterPosition, uint cNumChars, out IDebugDocumentContext ppsc);
    }
}