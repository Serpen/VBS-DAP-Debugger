'Option Explicit

sub my1()
    
    dim obj1 : set obj1 = new MYClass
    'obj1.m_sub

    tab = vbTab

    Set delegate = GetRef("caller")

    'set myobj2 = myobj
    'set ev = GetRef("myobj.handler")

    'myObj.Name = "Marco"
    'myObj13.Name = "Marco"

    dim zahl
    zahl = 4
    zahl = 4
    

    nul = null
    set no = nothing
    ne = empty

    arr = Array(12)

    cur = CCur(12)

    
    'msgbox name
    'i = CInt(myobj.Age)
    dbl = CDbl(18)
    sng = CSng(18)
    boo = CBool(true)
    byt = CByte(55)
    dat = CDate(now)
    lng = CLng(12)
    str = CStr("hallo")
    
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