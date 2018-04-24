Attribute VB_Name = "ProcessStatusCheck"
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

Public Function IsOpenProcess(ID_PROCc As Double) As Boolean
    Dim h As Long
    h = OpenProcess(&H1, True, ID_PROCc)
    If h <> 0 Then
      CloseHandle h
      IsOpenProcess = True
    Else
      IsOpenProcess = False
    End If
End Function
