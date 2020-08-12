VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form AkumKrm 
   Caption         =   "Lap Akumulasi Kiriman"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3600
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
   ScaleHeight     =   4125
   ScaleWidth      =   3600
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CR 
      Left            =   1560
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mingguan"
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   3375
      Begin VB.ComboBox Combo3 
         Height          =   345
         Left            =   1680
         TabIndex        =   2
         Top             =   720
         Width           =   1500
      End
      Begin VB.ComboBox Combo2 
         Height          =   345
         Left            =   1680
         TabIndex        =   1
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal Akhir"
         Height          =   345
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1500
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal Awal"
         Height          =   345
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1500
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Harian"
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3375
      Begin VB.ComboBox Combo1 
         Height          =   345
         Left            =   1680
         TabIndex        =   0
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal"
         Height          =   345
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1500
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Bulanan"
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   3375
      Begin VB.ComboBox Combo5 
         Height          =   345
         Left            =   1680
         TabIndex        =   4
         Top             =   720
         Width           =   1500
      End
      Begin VB.ComboBox Combo4 
         Height          =   345
         Left            =   1680
         TabIndex        =   3
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Bulan"
         Height          =   345
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tahun"
         Height          =   345
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1500
      End
   End
End
Attribute VB_Name = "AkumKrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
'On Error Resume Next
Call BukaDB
RSPengiriman.Open "Select Distinct Tanggal From pengiriman order By 1", Conn
RSPengiriman.Requery
Do Until RSPengiriman.EOF
    Combo1.AddItem Format(RSPengiriman!Tanggal, "DD-MMM-YYYY")
    Combo2.AddItem Format(RSPengiriman!Tanggal, "YYYY ,MM, DD")
    Combo3.AddItem Format(RSPengiriman!Tanggal, "YYYY ,MM, DD")
    RSPengiriman.MoveNext
Loop
Conn.Close

Call BukaDB
Dim RSTGL As New ADODB.Recordset
RSTGL.Open "select distinct month(Tanggal) as Bulan from pengiriman", Conn
Do While Not RSTGL.EOF
    Combo4.AddItem RSTGL!Bulan & Space(5) & MonthName(RSTGL!Bulan)
    RSTGL.MoveNext
Loop
Conn.Close

Call BukaDB
Dim RSTHN As New ADODB.Recordset
RSTHN.Open "select distinct year(Tanggal)  as Tahun from pengiriman", Conn
Do While Not RSTHN.EOF
    Combo5.AddItem RSTHN!Tahun
    RSTHN.MoveNext
Loop
Conn.Close

End Sub

Private Sub Combo1_Keypress(Keyascii As Integer)
If Combo1 = "" Or Keyascii = 27 Then Unload Me
End Sub

'Lap Harian
Private Sub Combo1_Click()
    CR.SelectionFormula = "Totext({Pengiriman.TanggalKrm})='" & Combo1 & "'"
    CR.ReportFileName = App.Path & "\Lap Kirim Harian.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub Combo2_Keypress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
End Sub

'Lap Mingguan (Tgl Antara)
Private Sub Combo3_Click()
    If Combo2 = "" Then
        MsgBox "TanggalKrm awal kosong", , "Informasi"
        Combo2.SetFocus
        Exit Sub
    End If
    CR.SelectionFormula = "{Pengiriman.TanggalKrm} in date (" & Combo2.Text & ") to date (" & Combo3.Text & ")"
    CR.ReportFileName = App.Path & "\Lap Kirim Mingguan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub Combo4_Keypress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
End Sub

'Lap Bulanan
Private Sub Combo5_Click()
    Call BukaDB
    RSPengiriman.Open "select * from Pengiriman where month(TanggalKrm)='" & Val(Combo4) & "' and year(TanggalKrm)='" & (Combo5) & "'", Conn
    If RSPengiriman.EOF Then
        MsgBox "Data tidak ditemukan"
        Exit Sub
        Combo4.SetFocus
    End If

    CR.SelectionFormula = "Month({Pengiriman.TanggalKrm})=" & Val(Combo4.Text) & " and Year({Pengiriman.TanggalKrm})=" & Val(Combo5.Text)
    CR.ReportFileName = App.Path & "\Lap Kirim Bulanan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub


