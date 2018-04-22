VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form EncodeForm 
   Caption         =   "Encode"
   ClientHeight    =   5985
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14760
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5985
   ScaleWidth      =   14760
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "ENCODE"
      Height          =   495
      Left            =   2160
      TabIndex        =   11
      Top             =   5400
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Browse"
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   4440
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   4215
      Left            =   6720
      ScaleHeight     =   4155
      ScaleWidth      =   5115
      TabIndex        =   6
      Top             =   1080
      Width           =   5175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox PasswordBox 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Text            =   "Password"
      Top             =   2520
      Width           =   3975
   End
   Begin VB.TextBox Text 
      Height          =   735
      Left            =   2160
      TabIndex        =   0
      Text            =   "Your text Here"
      Top             =   1200
      Width           =   3975
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   600
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label6 
      Caption         =   "Destination Image"
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Source Image"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "No File Selected!"
      Height          =   495
      Left            =   4080
      TabIndex        =   8
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "No File Selected!"
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   615
      Left            =   480
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Your Text Here"
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
End
Attribute VB_Name = "EncodeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

 On Error GoTo cancel
    CommonDialog1.Filter = "Image files (*.jpg, *.png, *.bmp, *.ico)|*.jpg;*.png;*.bmp;*.ico|PNG Files (*.png)|*.png|JPEG Files (*.jpg)|*.jpg|Bitmap Files (*.bmp)|*.bmp|Icon Files (*.ico)|*.ico"
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

Private Sub Command2_Click()
    
    If Label3.Caption = "No File Selected!" Then
        i = MsgBox("Selected Source Image First!", vbCritical, "Error")
    Else
        
        Path = Label3.Caption
        If Right(Path, 3) = "jpg" Then
            ext = "png"
        Else
            ext = Right(Path, 3)
        End If
       
        On Error GoTo cancel
        CommonDialog2.Filter = "Image files (*.png, *.bmp, *.ico)|*.png;*.bmp;*.ico|PNG Files (*.png)|*.png|Bitmap Files (*.bmp)|*.bmp|Icon Files (*.ico)|*.ico"
        CommonDialog2.DefaultExt = ext
        CommonDialog2.Flags = cdlOFNNoChangeDir Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist Or cdlOFNCreatePrompt
        CommonDialog2.ShowSave
        Path = CommonDialog2.Filename
        If Right(Path, 3) <> "bmp" And Right(Path, 3) <> "png" And Right(Path, 3) <> "ico" Then
            i = MsgBox("Inavlid File Type!", vbCritical, "Error")
            Label4.Caption = "No File Selected!"
        Else
            Label4.Caption = Path
        End If
    End If
    
cancel:
End Sub

Private Sub Command3_Click()
    
    If Label3.Caption <> "No File Selected!" And Label4.Caption <> "No File Selected!" Then

        Dim iFileNo, lineCount As Integer, ipString, opString As String
        iFileNo = FreeFile
        Open "ipe.txt" For Output As #iFileNo
        Print #iFileNo, Encrypt(Label3.Caption, 0)
        Print #iFileNo, Encrypt(Label4.Caption, 0)
        Print #iFileNo, Encrypt(PasswordBox, 79)
        Print #iFileNo, Encrypt(Text, 79)
        Close #iFileNo
        
        Dim val_exe As Double
        val_exe = Shell("encode.bat", vbHide)
    
        While IsOpenProcess(val_exe)
            'Wait for process to end
        Wend
    
        iFileNo = FreeFile
        Open "ope.txt" For Input As #iFileNo
        Line Input #iFileNo, opString
        If opString = "True" Then
            Line Input #iFileNo, opString
            MsgBox opString, vbOKOnly, "Results"
        Else
            Line Input #iFileNo, opString
            MsgBox opString, vbCritical, "Results"
        End If
        Close #iFileNo
        
        val_exe = Shell("removeEFiles.bat", vbHide)
    Else
        i = MsgBox("Image path(s) not set!", vbCritical, "ERROR")
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

Private Sub Label4_Click()
    If (Label4.Caption <> "No File Selected!") Then
        i = MsgBox(Label4.Caption, vbInformation, "Destination Image Path")
    Else
        i = MsgBox(Label4.Caption, vbCritical, "Destination Image Path")
    End If
End Sub

Private Sub PasswordBox_DblClick()
    PasswordBox = ""
    PasswordBox.PasswordChar = "*"
End Sub

Private Sub Text_DblClick()
    Text = ""
End Sub
