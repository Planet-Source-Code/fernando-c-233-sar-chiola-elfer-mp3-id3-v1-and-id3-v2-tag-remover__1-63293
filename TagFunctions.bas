Attribute VB_Name = "TagFunctions"
Public Function RemoveID31Tag(FileName As String)
If Not IsID31Present(FileName) Then Exit Function
Dim FileString As String
FileString = Space(FileLen(FileName) - 128)

nFic1 = FreeFile
Open FileName For Binary As nFic1
nFic2 = FreeFile
Open "NO-ID31-" + FileName For Binary As nFic2
Get nFic1, 1, FileString
Put nFic2, , FileString
Close nFic1
Close nFic2

End Function
Public Function RemoveID32Tag(FileName As String)
If Not IsID32Present(FileName) Then Exit Function
Dim FileString As String
FileString = Space(FileLen(FileName) - ID32Size(FileName))

nFic1 = FreeFile
Open FileName For Binary As nFic1
nFic2 = FreeFile
Open "NO-ID32-" + FileName For Binary As nFic2
Get nFic1, ID32Size(FileName) + 1, FileString
Put nFic2, , FileString
Close nFic1
Close nFic2

End Function
Public Function RemoveLyricTag(FPath As String, FRecurse As Boolean)

End Function

Public Function IsID31Present(FileName As String) As Boolean
Dim ID31 As String * 3 '''Space$(3)

nFic = FreeFile
Open FileName For Binary As nFic
Get nFic, FileLen(FileName) - 127, ID31
Close nFic

If ID31 = "TAG" Then IsID31Present = True Else IsID31Present = False

End Function

Public Function IsID32Present(FileName As String) As Boolean
Dim ID32 As String * 3 '''Space$(3)

nFic = FreeFile
Open FileName For Binary As nFic
Get nFic, 1, ID32
Close nFic

If ID32 = "ID3" Then IsID32Present = True Else IsID32Present = False

End Function

Public Function ID32Size(FileName As String) As Long
If Not IsID32Present(FileName) Then Exit Function
Dim ID32FourBytes As String * 4 '''Space$(4)
Dim ByteX As String

nFic = FreeFile
Open FileName For Binary As nFic
Get nFic, 7, ID32FourBytes
Close nFic

ID32Size = 1
For x = 1 To 4
ByteX = Mid$(ID32FourBytes, x, 1)
If Asc(ByteX) <> 0 Then ID32Size = ID32Size * Asc(ByteX)
Next

End Function

Public Function ID31Size(FileName As String) As Byte
If Not IsID31Present(FileName) Then Exit Function
ID31Size = 128
End Function

Public Function MP3Size(FileName As String) As Long
MP3Size = FileLen(FileName) - ID32Size(FileName) - ID31Size(FileName)
End Function

