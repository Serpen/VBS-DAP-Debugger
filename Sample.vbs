'Option Explicit

sub my1()
    
    dim obj1 : set obj1 = new MYClass
    'obj1.m_sub

    tab = vbTab

    Set delegate = GetRef("caller") 

    nul = null
    set no = nothing
    ne = empty

    arr = Array(4)

    boo = CBool(true)
    byt = CByte(55)
    cur = CCur(12.789)
    dat = CDate(now)
    dbl = CDbl(18)
    lng = CLng(12)
    inte = CInt(12)
    sng = CSng(18)
    str = CStr("hallo")

    Debug.Write "debugwrite"

    Call MsgBox(cur, vbYesNoCancel)
    
    set fso = CreateObject("Scripting.FileSystemObject")
    Set dic = CreateObject("Scripting.Dictionary")

    Call dic.Add("index","value")

    vt = VarType(lng)
    vtname = TypeName(lng)

    Ok = vbOK

    f = false
    t = true

    stop




    
    
end sub

sub caller()
    my1
end sub

class MyClass
    private m_private
    public m_public

    public sub m_sub() 
        
    End sub

    function m_function() : End function
end class

caller