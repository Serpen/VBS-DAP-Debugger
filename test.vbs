Option Explicit
dim fso
msgbox "vor fso"
set fso = CreateObject("Scripting.FileSystemObject")
msgbox "vor stop"
stop
msgbox "nach stop"