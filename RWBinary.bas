Attribute VB_Name = "RWBinary"
Public Sub WriteBinary()
Dim fnum As Integer
file_name = "c:\minu.bin"
On Error Resume Next
Kill file_name
On Error GoTo 0

' Save the file.
fnum = FreeFile
Open file_name For Binary As #fnum
Put #fnum, 1, bHuellaMinuciasCAP
Close fnum
End Sub
Public Sub ReadBinary()
Dim file_name As String
Dim fnum As Integer
Dim file_length As Long
file_name = "c:\minu.bin"
file_length = FileLen(file_name)
ReDim bHuellaMinucias(1 To file_length)
fnum = FreeFile
Open file_name For Binary As #fnum
Get #fnum, 1, bHuellaMinucias
Close fnum

End Sub
