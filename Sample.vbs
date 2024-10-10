'Option Explicit

sub my1()
    'stop
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
    dat = CDate(#11/11/2022 10:01:17 PM#)
    dbl = CDbl(Sqr(18.4))
    lng = CLng(12)
    inte = CInt(12)
    sng = CSng(Sqr(18.4))
    str = CStr("hallo")

    dim i : i = 0
    Do
        i = i+1
    Loop While i < 1000

    Debug.Write "debugwrite"

    'stop
    'Call MsgBox(cur, vbYesNoCancel)
    
    set fso = CreateObject("Scripting.FileSystemObject")
    Set dic = CreateObject("Scripting.Dictionary")

    Call dic.Add("index","value")

    vt = VarType(lng)
    vtname = TypeName(lng)

    Ok = vbOK

    f = false
    t = True

    stop
    
    const myval = "const"

    msgbox 1


    
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

    Public Property Get PropertyName
        PropertyName = m_private
    End Property
    Public Property Let PropertyName(Value)
        m_private = Value
    End Property
    
    Private Sub Class_Initialize()
        m_private = 66
        set m_public = new MyClass
    End Sub

end class

caller