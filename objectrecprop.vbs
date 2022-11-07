
dim obj
dim zahl
dim fso

zahl = 4
set fso = CreateObject("Scripting.FileSystemObject")

set obj = new objectrecprop
Stop

Class objectrecprop

    Private Sub Class_Initialize()
        arr = Array(5)
        myprivate = 12
        m_PropertyName = 18
    End Sub

    Private Sub Class_Terminate()
        
    End Sub

    Public Property Get PropertyName
        PropertyName = m_PropertyName
    End Property
    Public Property Let PropertyName(Value)
        m_PropertyName = Value
    End Property

    private myprivate

    public [SpecialName]

    ' public const RAD1 = 12

    public arr 





End Class ' objectrecprop