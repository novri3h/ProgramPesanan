VERSION 5.00
Begin VB.Form Tampilkan 
   BackColor       =   &H80000009&
   Caption         =   "ESC / Enter = Tutup"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4275
   BeginProperty Font 
      Name            =   "Century"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   4275
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Tampilkan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(Keyascii As Integer)
If Keyascii = 27 Or Keyascii = 13 Then Unload Me
End Sub

