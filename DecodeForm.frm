VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form DecodeForm 
   Caption         =   "Decode"
   ClientHeight    =   5985
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14760
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5985
   ScaleWidth      =   14760
   WindowState     =   2  'Maximized
   Begin VB.TextBox PasswordBox 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Text            =   "Password"
      Top             =   3120
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   4215
      Left            =   6720
      ScaleHeight     =   4155
      ScaleWidth      =   5115
      TabIndex        =   1
      Top             =   1080
      Width           =   5175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "DECODE"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   4800
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "No File Selected!"
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Source Image"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
   End
End
Attribute VB_Name = "DecodeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

 On Error GoTo cancel
    CommonDialog1.Filter = "Image files (*.png, *.bmp, *.ico)|*.png;*.bmp;*.ico|PNG Files (*.png)|*.png|Bitmap Files (*.bmp)|*.bmp|Icon Files (*.ico)|*.ico"
    CommonDialog1.Flags = cdlOFNFileMustExist Or cdlOFNNoChangeDir
    CommonDialog1.ShowOpen
    Path = CommonDialog1.Filename
    
    If Right(Path, 3) <> "jpg" And Right(Path, 3) <> "bmp" And Right(Path, 3) <> "png" And Right(Path, 3) <> "ico" Then
        i = MsgBox("Inavlid File Type!", vbCritical, "Error")
        Label3.Caption = "No File Selected!"
    Else
        Label3.Caption = Path
    End If
    
cancel:
End Sub
Private Sub Command3_Click()

    If Label3.Caption <> "No File Selected!" Then

        Dim iFileNo, lineCount As Integer, line, Text As String
        iFileNo = FreeFile
        Open "ipd.txt" For Output As #iFileNo
        Print #iFileNo, Encrypt(Label3.Caption, 0)
        Print #iFileNo, Encrypt(PasswordBox, 79)
        Close #iFileNo
        
        Dim val_exe As Double
        val_exe = Shell("decode.bat", vbHide)
    
        While IsOpenProcess(val_exe)
            'Wait for process to end
        Wend
    
        iFileNo = FreeFile
        Open "opd.txt" For Input As #iFileNo
        Line Input #iFileNo, line
        If line = "True" Then
            Do Until EOF(iFileNo)
                Line Input #iFileNo, line
                Text = Text + line + vbCrLf
            Loop
            i = MsgBox(Encrypt(Left(Text, Len(Text) - 2), -79), vbOKOnly, "Hidden Text")
        Else
            Line Input #iFileNo, line
            MsgBox line, vbCritical, "Results"
        End If
        Close #iFileNo
        
        val_exe = Shell("removeDFiles.bat", vbHide)
    Else
        i = MsgBox("Image path not set!", vbCritical, "ERROR")
    End If

End Sub

Private Sub Label3_Change()
    If Label3.Caption <> "No File Selected!" Then
        Picture1.Picture = Nothing
        Picture1.Cls
        Picture1.Picture = Picture()
        Picture1.Picture = LoadPictureEx(Label3.Caption)
        Picture1.PaintPicture Picture1.Picture, 0, 0, 5175, 4215, opcode = vbSrcCopy
        Picture1.AutoRedraw = False
        Picture1.ScaleMode = vbPixels
    End If
End Sub

Private Sub Label3_Click()
    If (Label3.Caption <> "No File Selected!") Then
        i = MsgBox(Label3.Caption, vbInformation, "Source Image Path")
    Else
        i = MsgBox(Label3.Caption, vbCritical, "Source Image Path")
    End If
End Sub

Private Sub PasswordBox_DblClick()
    PasswordBox = ""
    PasswordBox.PasswordChar = "*"
End Sub

