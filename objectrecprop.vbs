
dim fso
dim zahl
dim obj

zahl = 4
set fso = CreateObject("Scripting.FileSystemObject")

set obj = new objectrecprop
Stop

Class objectrecprop

    Private Sub Class_Initialize()
        
    End Sub

    Private Sub Class_Terminate()
        
    End Sub

    Public Property Get PropertyName
        PropertyName = m_PropertyName
    End Property
    Public Property Let PropertyName(Value)
        m_PropertyName = Value
    End Property


End Class ' objectrecprop