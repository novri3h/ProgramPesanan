VERSION 5.00
Begin VB.Form Kurir 
   Caption         =   "Kurir"
   ClientHeight    =   1860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4830
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
   ScaleHeight     =   1860
   ScaleWidth      =   4830
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton CmdTutup 
         Caption         =   "&Tutup"
         Height          =   350
         Left            =   3360
         TabIndex        =   3
         Top             =   1080
         Width           =   1000
      End
      Begin VB.CommandButton CmdHapus 
         Caption         =   "&Hapus"
         Height          =   350
         Left            =   2280
         TabIndex        =   2
         Top             =   1080
         Width           =   1000
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "&Edit"
         Height          =   350
         Left            =   1200
         TabIndex        =   1
         Top             =   1080
         Width           =   1000
      End
      Begin VB.CommandButton CmdInput 
         Caption         =   "&Input"
         Height          =   350
         Left            =   120
         TabIndex        =   0
         Top             =   1080
         Width           =   1000
      End
      Begin VB.TextBox Text2 
         Height          =   350
         Left            =   1200
         TabIndex        =   5
         Top             =   600
         Width           =   3200
      End
      Begin VB.TextBox Text1 
         Height          =   350
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nama"
         Height          =   345
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Kode"
         Height          =   345
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1005
      End
   End
End
Attribute VB_Name = "Kurir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Form_Load()
    Call BukaDB
    Text1.MaxLength = 5
    Text2.MaxLength = 30
    KondisiAwal
End Sub

Function CariData()
    Call BukaDB
    RSKurir.Open "Select * From Kurir where KodeKrr='" & Text1 & "'", Conn
End Function

Private Sub KosongkanText()
    Text1 = ""
    Text2 = ""
End Sub

Private Sub SiapIsi()
    Text1.Enabled = True
    Text2.Enabled = True
End Sub

Private Sub TidakSiapIsi()
    Text1.Enabled = False
    Text2.Enabled = False
End Sub

Private Sub KondisiAwal()
    KosongkanText
    TidakSiapIsi
    CmdInput.Caption = "&Input"
    CmdEdit.Caption = "&Edit"
    CmdHapus.Caption = "&Hapus"
    CmdTutup.Caption = "&Tutup"
    CmdInput.Enabled = True
    CmdEdit.Enabled = True
    CmdHapus.Enabled = True
End Sub

Private Sub TampilkanData()
    With RSKurir
        If Not RSKurir.EOF Then
            Text2 = RSKurir!NamaKrr
        End If
    End With
End Sub

Private Sub CmdInput_Click()
    If CmdInput.Caption = "&Input" Then
        CmdInput.Caption = "&Simpan"
        CmdEdit.Enabled = False
        CmdHapus.Enabled = False
        CmdTutup.Caption = "&Batal"
        SiapIsi
        KosongkanText
        Text1.SetFocus
    Else
        If Text1 = "" Or Text2 = "" Then
            MsgBox "Data Belum Lengkap...!"
        Else
            Dim SQLTambah As String
            SQLTambah = "Insert Into Kurir (KodeKrr,NamaKrr) values ('" & Text1 & "','" & Text2 & "')"
            Conn.Execute SQLTambah
            KondisiAwal
        End If
    End If
End Sub

Private Sub CmdEdit_Click()
    If CmdEdit.Caption = "&Edit" Then
        CmdInput.Enabled = False
        CmdEdit.Caption = "&Simpan"
        CmdHapus.Enabled = False
        CmdTutup.Caption = "&Batal"
        SiapIsi
        Text1.SetFocus
    Else
        If Text2 = "" Then
            MsgBox "Masih Ada Data Yang Kosong"
        Else
            Dim SQLEdit As String
            SQLEdit = "Update Kurir Set NamaKrr= '" & Text2 & "' where KodeKrr='" & Text1 & "'"
            Conn.Execute SQLEdit
            KondisiAwal
        End If
    End If
End Sub

Private Sub CmdHapus_Click()
    If CmdHapus.Caption = "&Hapus" Then
        CmdInput.Enabled = False
        CmdEdit.Enabled = False
        CmdTutup.Caption = "&Batal"
        KosongkanText
        SiapIsi
        Text1.SetFocus
    End If
End Sub

Private Sub CmdTutup_Click()
    Select Case CmdTutup.Caption
        Case "&Tutup"
            Unload Me
        Case "&Batal"
            TidakSiapIsi
            KondisiAwal
    End Select
End Sub

Private Sub Text1_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Len(Text1) < 5 Then
        MsgBox "Kode Harus 5 Digit"
        Text1.SetFocus
    Else
        Text2.SetFocus
    End If

    If CmdInput.Caption = "&Simpan" Then
        Call CariData
            If Not RSKurir.EOF Then
                TampilkanData
                MsgBox "Kode Kurir Sudah Ada"
                KosongkanText
                Text1.SetFocus
            Else
                Text2.SetFocus
            End If
    End If
    
    If CmdEdit.Caption = "&Simpan" Then
        Call CariData
            If Not RSKurir.EOF Then
                TampilkanData
                Text1.Enabled = False
                Text2.SetFocus
            Else
                MsgBox "Kode Kurir Tidak Ada"
                Text1 = ""
                Text1.SetFocus
            End If
    End If
    
    If CmdHapus.Enabled = True Then
        Call CariData
            If Not RSKurir.EOF Then
                TampilkanData
                Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
                If Pesan = vbYes Then
                    Dim SQLHapus As String
                    SQLHapus = "Delete From Kurir where kodeKrr= '" & Text1 & "'"
                    Conn.Execute SQLHapus
                    KondisiAwal
                Else
                    KondisiAwal
                    CmdHapus.SetFocus
                End If
            Else
                MsgBox "Data Tidak ditemukan"
                Text1.SetFocus
            End If
    End If
End If
End Sub

Private Sub Text2_keypress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then
        If CmdInput.Enabled = True Then
            CmdInput.SetFocus
        ElseIf CmdEdit.Enabled = True Then
            CmdEdit.SetFocus
        End If
    End If
End Sub




