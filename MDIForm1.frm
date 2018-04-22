VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Steganography"
   ClientHeight    =   6570
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12780
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu Encode 
      Caption         =   "&Encode"
   End
   Begin VB.Menu Decode 
      Caption         =   "&Decode"
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
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
Private Sub Decode_Click()
    If EncodeForm.Visible = True Then
        EncodeForm.Hide
    End If
    DecodeForm.Show
End Sub

Private Sub Encode_Click()
    If DecodeForm.Visible = True Then
        DecodeForm.Hide
    End If
    EncodeForm.Show
End Sub

