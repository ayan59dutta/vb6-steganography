VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Steganography"
   ClientHeight    =   6570
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   12780
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":1CCA0
   Begin VB.Menu Encode 
      Caption         =   "&Encode"
   End
   Begin VB.Menu Decode 
      Caption         =   "&Decode"
   End
   Begin VB.Menu About 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub About_Click()
    If EncodeForm.Visible = True Then
        EncodeForm.Hide
    End If
    If DecodeForm.Visible = True Then
        DecodeForm.Hide
    End If
    AboutForm.WindowState = 2
    AboutForm.Show
End Sub

Private Sub Decode_Click()
    If EncodeForm.Visible = True Then
        EncodeForm.Hide
    End If
    If AboutForm.Visible = True Then
        EncodeForm.Hide
    End If
    DecodeForm.Show
End Sub

Private Sub Encode_Click()
    If DecodeForm.Visible = True Then
        DecodeForm.Hide
    End If
    If AboutForm.Visible = True Then
        EncodeForm.Hide
    End If
    EncodeForm.Show
End Sub

