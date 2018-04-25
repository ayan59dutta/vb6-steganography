VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form EncodeForm 
   Caption         =   "Encode"
   ClientHeight    =   5985
   ClientLeft      =   1830
   ClientTop       =   2475
   ClientWidth     =   14760
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "EncodeForm.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "EncodeForm.frx":1CCA0
   ScaleHeight     =   5985
   ScaleWidth      =   14760
   WindowState     =   2  'Maximized
   Begin VB.CommandButton EncodeCmd 
      Caption         =   "&ENCODE"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   11
      Top             =   5280
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton SaveCmd 
      Caption         =   "Browse"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   4320
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   6720
      ScaleHeight     =   4155
      ScaleWidth      =   5115
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.CommandButton OpenCmd 
      Caption         =   "Browse"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox PasswordBox 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Text            =   "Password"
      Top             =   2400
      Width           =   3975
   End
   Begin VB.TextBox Text 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   0
      Text            =   "Enter Your Text Here"
      Top             =   1080
      Width           =   3975
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   600
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Destination Image"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   480
      TabIndex        =   10
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Source Image"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label SavePath 
      BackStyle       =   0  'Transparent
      Caption         =   "No File Selected!"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   4080
      TabIndex        =   8
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label OpenPath 
      BackStyle       =   0  'Transparent
      Caption         =   "No File Selected!"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   4080
      TabIndex        =   5
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   480
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Text To Be Hidden"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
End
Attribute VB_Name = "EncodeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Picture1.Cls
End Sub

Private Sub OpenCmd_Click()
    Dim Path As String, i As Integer
    On Error GoTo cancel
    CommonDialog1.Filter = "Image files (*.jpg, *.png, *.bmp, *.ico)|*.jpg;*.png;*.bmp;*.ico|PNG Files (*.png)|*.png|JPEG Files (*.jpg)|*.jpg|Bitmap Files (*.bmp)|*.bmp|Icon Files (*.ico)|*.ico"
    CommonDialog1.Flags = cdlOFNFileMustExist Or cdlOFNNoChangeDir
    CommonDialog1.ShowOpen
    Path = CommonDialog1.Filename
    If Right(Path, 3) <> "jpg" And Right(Path, 3) <> "bmp" And Right(Path, 3) <> "png" And Right(Path, 3) <> "ico" Then
        i = MsgBox("Inavlid File Type!", vbCritical, "Error")
        OpenPath.Caption = "No File Selected!"
    Else
        OpenPath.Caption = Path
        SaveCmd.Enabled = True
    End If
cancel:
End Sub

Private Sub SaveCmd_Click()
    Dim i As Integer, Path, ext As String
    If OpenPath.Caption = "No File Selected!" Then
        i = MsgBox("Selected Source Image First!", vbCritical, "Error")
    Else
        Path = OpenPath.Caption
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
            SavePath.Caption = "No File Selected!"
        Else
            SavePath.Caption = Path
            EncodeCmd.Visible = True
            EncodeCmd.SetFocus
        End If
    End If
cancel:
End Sub

Private Sub EncodeCmd_Click()
    Dim i As Integer
    If OpenPath.Caption <> "No File Selected!" And SavePath.Caption <> "No File Selected!" Then
        Dim iFileNo, lineCount As Integer, val_exe As Double, opString As String
        iFileNo = FreeFile
        Open "ipe.txt" For Output As #iFileNo
        Print #iFileNo, OpenPath.Caption
        Print #iFileNo, SavePath.Caption
        Print #iFileNo, PasswordBox
        Print #iFileNo, Text
        Close #iFileNo
        val_exe = Shell("encode.bat", vbHide)
        While IsOpenProcess(val_exe)
            'Wait for process to end
        Wend
        iFileNo = FreeFile
        Open "ope.txt" For Input As #iFileNo
        Line Input #iFileNo, opString
        If opString = "True" Then
            Line Input #iFileNo, opString
            Close #iFileNo
            Kill ("ope.txt")
            MsgBox opString, vbOKOnly, "Results"
            OpenPath.Caption = SavePath.Caption
            MsgBox "The encoded image is now displayed in the Picture Box.", vbOKOnly, "Results"
        Else
            Line Input #iFileNo, opString
            Close #iFileNo
            Kill ("ope.txt")
            MsgBox opString, vbCritical, "Results"
        End If
    Else
        i = MsgBox("Image path(s) not set!", vbCritical, "ERROR")
    End If
End Sub

Private Sub OpenPath_Change()
    If OpenPath.Caption <> "No File Selected!" Then
        Picture1.AutoRedraw = True
        Picture1.Visible = True
        Picture1.Picture = LoadPictureEx(OpenPath.Caption)
        Picture1.PaintPicture Picture1.Picture, 0, 0, 5175, 4215
    Else
        Picture1.Visible = False
    End If
End Sub

Private Sub OpenPath_Click()
    Dim i As Integer
    If (OpenPath.Caption <> "No File Selected!") Then
        i = MsgBox(OpenPath.Caption, vbInformation, "Source Image Path")
    Else
        i = MsgBox(OpenPath.Caption, vbCritical, "Source Image Path")
    End If
End Sub

Private Sub SavePath_Click()
    Dim i As Integer
    If (SavePath.Caption <> "No File Selected!") Then
        i = MsgBox(SavePath.Caption, vbInformation, "Destination Image Path")
    Else
        i = MsgBox(SavePath.Caption, vbCritical, "Destination Image Path")
    End If
End Sub

Private Sub PasswordBox_DblClick()
    PasswordBox = ""
    PasswordBox.PasswordChar = "*"
    PasswordBox.FontSize = 16
End Sub

Private Sub Text_DblClick()
    Text = ""
End Sub
