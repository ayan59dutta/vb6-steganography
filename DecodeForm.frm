VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form DecodeForm 
   Caption         =   "Decode"
   ClientHeight    =   5985
   ClientLeft      =   2115
   ClientTop       =   2775
   ClientWidth     =   14760
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "DecodeForm.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "DecodeForm.frx":1CCA0
   ScaleHeight     =   5985
   ScaleWidth      =   14760
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox PasswordBox 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Text            =   "Password"
      Top             =   3120
      Width           =   3975
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
      Left            =   2040
      TabIndex        =   2
      Top             =   1560
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
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.CommandButton DecodeCmd 
      Caption         =   "&DECODE"
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
      Left            =   2040
      TabIndex        =   0
      Top             =   4800
      Visible         =   0   'False
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
      Left            =   360
      TabIndex        =   6
      Top             =   3240
      Width           =   1095
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
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Top             =   1680
      Width           =   2055
      WordWrap        =   -1  'True
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
Option Explicit
Private Sub OpenCmd_Click()
    Dim path As String, i As Integer
    On Error GoTo cancel
    CommonDialog1.Filter = "Image files (*.png, *.bmp, *.ico)|*.png;*.bmp;*.ico|PNG Files (*.png)|*.png|Bitmap Files (*.bmp)|*.bmp|Icon Files (*.ico)|*.ico"
    CommonDialog1.Flags = cdlOFNFileMustExist Or cdlOFNNoChangeDir
    CommonDialog1.ShowOpen
    path = CommonDialog1.Filename
    If Right(path, 3) <> "jpg" And Right(path, 3) <> "bmp" And Right(path, 3) <> "png" And Right(path, 3) <> "ico" Then
        i = MsgBox("Inavlid File Type!", vbCritical, "Error")
        OpenPath.Caption = "No File Selected!"
    Else
        OpenPath.Caption = path
        DecodeCmd.Visible = True
        DecodeCmd.SetFocus
    End If
cancel:
End Sub
Private Sub DecodeCmd_Click()
    If OpenPath.Caption <> "No File Selected!" Then
        Dim iFileNo, lineCount, i As Integer, valExe As Double, line, Text As String
        iFileNo = FreeFile
        Open "ipd.txt" For Output As #iFileNo
        Print #iFileNo, OpenPath.Caption
        Print #iFileNo, PasswordBox
        Close #iFileNo
        valExe = Shell("decode.bat", vbHide)
        While IsOpenProcess(valExe)
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
            Close #iFileNo
            Kill ("opd.txt")
            i = MsgBox(Left(Text, Len(Text) - 2), vbOKOnly, "Hidden Text")
        Else
            Line Input #iFileNo, line
            Close #iFileNo
            Kill ("ipd.txt")
            MsgBox line, vbCritical, "Results"
        End If
    Else
        i = MsgBox("Image path not set!", vbCritical, "ERROR")
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

Private Sub PasswordBox_DblClick()
    PasswordBox = ""
    PasswordBox.PasswordChar = "*"
    PasswordBox.FontSize = 16
End Sub
