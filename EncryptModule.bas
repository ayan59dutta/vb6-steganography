Attribute VB_Name = "Module1"
Public Function Encrypt(PlainText As String, key As Integer) As String
    
    Dim i As Integer, CipherText As String
    CipherText = ""
    
    For i = 1 To Len(PlainText)
        If ((Asc(Mid(PlainText, i, 1)) + key) Mod 256) >= 0 Then
            CipherText = CipherText + Chr((Asc(Mid(PlainText, i, 1)) + key) Mod 256)
        Else
            CipherText = CipherText + Chr(((Asc(Mid(PlainText, i, 1)) + key) Mod 256) + 256)
        End If
    Next i
    
    Encrypt = CipherText
    
End Function
