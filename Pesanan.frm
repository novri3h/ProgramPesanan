VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Pesanan 
   Caption         =   "Data Pesanan"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9975
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
   LockControls    =   -1  'True
   ScaleHeight     =   5190
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   4785
      Left            =   7080
      TabIndex        =   5
      Top             =   120
      Width           =   2775
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Pesanan.frx":0000
      Height          =   1635
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   2884
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "Nomor"
         Caption         =   "Nomor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Kode"
         Caption         =   "Kode"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Nama"
         Caption         =   "Nama"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Harga"
         Caption         =   "Harga"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Jumlah"
         Caption         =   "Jumlah"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Total"
         Caption         =   "Total"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   645,165
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2505,26
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   615,118
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   915,024
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identitas Pemesan"
      Height          =   2055
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   4335
      Begin VB.ComboBox Combo1 
         Height          =   345
         Left            =   2280
         TabIndex        =   0
         Text            =   "Combo1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Teleponksm 
         Height          =   350
         Left            =   1200
         TabIndex        =   1
         Top             =   720
         Width           =   3000
      End
      Begin VB.TextBox Alamatksm 
         Height          =   350
         Left            =   1200
         TabIndex        =   3
         Top             =   1440
         Width           =   3000
      End
      Begin VB.TextBox Namaksm 
         Height          =   350
         Left            =   1200
         TabIndex        =   2
         Top             =   1080
         Width           =   3000
      End
      Begin VB.TextBox Nomorksm 
         Height          =   350
         Left            =   1200
         TabIndex        =   7
         Top             =   360
         Width           =   1000
      End
      Begin VB.Label Label20 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Telepon"
         Height          =   345
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label19 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Alamat"
         Height          =   345
         Left            =   120
         TabIndex        =   26
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label Label18 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nama"
         Height          =   345
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label Label17 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nomor"
         Height          =   345
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   1005
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2160
      Top             =   4440
   End
   Begin VB.TextBox DP 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   5760
      TabIndex        =   6
      Top             =   4320
      Width           =   1100
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   350
      Left            =   1800
      TabIndex        =   10
      Top             =   3960
      Width           =   850
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   350
      Left            =   960
      TabIndex        =   9
      Top             =   3960
      Width           =   850
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   350
      Left            =   120
      TabIndex        =   8
      Top             =   3960
      Width           =   850
   End
   Begin MSAdodcLib.Adodc DT 
      Height          =   405
      Left            =   120
      Top             =   4440
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   714
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Transaksi"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker TglMintakrm 
      Height          =   300
      Left            =   5640
      TabIndex        =   30
      Top             =   240
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   94830593
      CurrentDate     =   39312
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kirimkan Tgl"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4560
      TabIndex        =   31
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Item"
      Height          =   345
      Left            =   2760
      TabIndex        =   29
      Top             =   3960
      Width           =   600
   End
   Begin VB.Label JmlItem 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   3480
      TabIndex        =   28
      Top             =   3960
      Width           =   600
   End
   Begin VB.Label Sisa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   5760
      TabIndex        =   22
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Total 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   5760
      TabIndex        =   21
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Sisa"
      Height          =   300
      Left            =   4680
      TabIndex        =   20
      Top             =   4680
      Width           =   1005
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Uang Muka"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4680
      TabIndex        =   19
      Top             =   4320
      Width           =   1005
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total"
      Height          =   300
      Left            =   4680
      TabIndex        =   18
      Top             =   3960
      Width           =   1005
   End
   Begin VB.Label Tanggal 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   5640
      TabIndex        =   17
      Top             =   1320
      Width           =   1245
   End
   Begin VB.Label Namaksr 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   5640
      TabIndex        =   16
      Top             =   960
      Width           =   1245
   End
   Begin VB.Label NomorPsn 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   5640
      TabIndex        =   15
      Top             =   600
      Width           =   1245
   End
   Begin VB.Label Kodeksr 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   2760
      TabIndex        =   14
      Top             =   4440
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal"
      Height          =   300
      Left            =   4560
      TabIndex        =   13
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kasir"
      Height          =   300
      Left            =   4560
      TabIndex        =   12
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nomor"
      Height          =   300
      Left            =   4560
      TabIndex        =   11
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "Pesanan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub form_Activate()
DT.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOPesanan.mdb"
DT.RecordSource = "Transaksi"
Set DataGrid1.DataSource = DT
DataGrid1.Refresh
   
If Kodeksr = "" Then
    MsgBox "Kasir tidak terdeteksi"
    Login.Show
    Exit Sub
End If

Call BukaDB
RSBarang.Open "Barang", Conn
List1.Clear
Do Until RSBarang.EOF
    List1.AddItem RSBarang!KodeBrg & vbTab & RSBarang!NamaBrg
    RSBarang.MoveNext
Loop

RSKonsumen.Open "Konsumen", Conn
Combo1.Clear
Do Until RSKonsumen.EOF
    Combo1.AddItem RSKonsumen!NomorKsm
    RSKonsumen.MoveNext
Loop

Call AutoPsn
Call AutoKsm
Call Tabel_Kosong
DT.Recordset.MoveFirst
Tanggal = Format(Date, "dd-mm-yyyy")
TglMintakrm.Value = Date
NomorKsm.Enabled = False
CmdSimpan.Enabled = False
End Sub

Private Sub Form_Load()
    Kodeksr = Login.TxtKodeKsr
    Namaksr = Login.TxtNamaKsr
    DataGrid1.Col = 1
    CmdSimpan.Enabled = False
End Sub

Private Sub Timer1_Timer()
    Jam = Time$
End Sub

Private Sub AutoPsn()
Call BukaDB
RSPesanan.Open ("select * from pesanan Where NomorPsn In(Select Max(NomorPsn)From Pesanan)Order By NomorPsn Desc"), Conn
RSPesanan.Requery
    Dim Urutan As String * 10
    Dim Hitung As Long
    With RSPesanan
        If .EOF Then
            Urutan = "P" + Format(Date, "yymmdd") + "001"
            NomorPsn = Urutan
        Else
            If Mid(!NomorPsn, 2, 6) <> Format(Date, "yymmdd") Then
             Urutan = "P" + Format(Date, "yymmdd") + "001"
            Else
                Hitung = Right(!NomorPsn, 9) + 1
                Urutan = "P" + Format(Date, "yymmdd") + Right("000" & Hitung, 3)
            End If
        End If
        NomorPsn = Urutan
    End With
End Sub

Private Sub AutoKsm()
Call BukaDB
RSKonsumen.Open ("select * from Konsumen Where NomorKsm In(Select Max(NomorKsm)From Konsumen)Order By NomorKsm Desc"), Conn
RSKonsumen.Requery
    Dim Urutan As String * 5
    Dim Hitung As Long
    With RSKonsumen
        If .EOF Then
            Urutan = "KSM01"
            NomorKsm = Urutan
        Else
            Hitung = Right(!NomorKsm, 2) + 1
            Urutan = "KSM" + Right("00" & Hitung, 2)
        End If
        NomorKsm = Urutan
    End With
End Sub

Private Sub Nomorksm_Change()
Call BukaDB
RSKonsumen.Open "Select * from konsumen where nomorksm='" & NomorKsm & "'", Conn
RSKonsumen.Requery
If Not RSKonsumen.EOF Then
    NamaKsm = RSKonsumen!NamaKsm
    AlamatKsm = RSKonsumen!AlamatKsm
    TeleponKsm = RSKonsumen!TeleponKsm
End If
End Sub

Private Sub teleponksm_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    Call BukaDB
    RSKonsumen.Open "Select * from konsumen where teleponksm='" & TeleponKsm & "'", Conn
    RSKonsumen.Requery
    If Not RSKonsumen.EOF Then
        NomorKsm = RSKonsumen!NomorKsm
        NamaKsm = RSKonsumen!NamaKsm
        AlamatKsm = RSKonsumen!AlamatKsm
        TeleponKsm = RSKonsumen!TeleponKsm
        List1.SetFocus
    Else
        NamaKsm.SetFocus
    End If
End If
If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub Namaksm_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    Call BukaDB
    RSKonsumen.Open "Select * from konsumen where namaksm='" & NamaKsm & "'", Conn
    RSKonsumen.Requery
    If Not RSKonsumen.EOF Then
        NomorKsm = RSKonsumen!NomorKsm
        AlamatKsm = RSKonsumen!AlamatKsm
        TeleponKsm = RSKonsumen!TeleponKsm
    End If
    AlamatKsm.SetFocus
End If
End Sub

Private Sub alamatksm_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    Call BukaDB
    RSKonsumen.Open "Select * from konsumen where alamatksm='" & AlamatKsm & "'", Conn
    RSKonsumen.Requery
    If Not RSKonsumen.EOF Then
        NomorKsm = RSKonsumen!NomorKsm
        NamaKsm = RSKonsumen!NamaKsm
        TeleponKsm = RSKonsumen!TeleponKsm
    End If
    DataGrid1.SetFocus
End If
End Sub

Private Sub Combo1_Keypress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Combo1 = "" Then
        Call AutoKsm
        MsgBox "silakan isi data konsumen baru"
        Kosongksm
        TeleponKsm.SetFocus
        Exit Sub
    Else
        DataGrid1.SetFocus
    End If
End If
If Keyascii = 27 Then
    Combo1 = ""
    Call AutoKsm
    MsgBox "silakan isi data konsumen baru"
    Kosongksm
    TeleponKsm.SetFocus
    Exit Sub
End If
End Sub

Private Sub Combo1_Click()
    Call BukaDB
    RSKonsumen.Open "Select * from Konsumen where Nomorksm='" & Combo1 & "'", Conn
    If Not RSKonsumen.EOF Then
        NomorKsm = RSKonsumen!NomorKsm
    End If
    Conn.Close
End Sub

Function Tabel_Kosong()
    DT.Recordset.MoveFirst
    Do While Not DT.Recordset.EOF
        DT.Recordset.Delete
        DT.Recordset.MoveNext
    Loop
    For i = 1 To 1
        DT.Recordset.AddNew
        DT.Recordset!Nomor = i
        DT.Recordset.Update
    Next i
    DataGrid1.Col = 1
End Function

Function Tambah_Baris()
    For i = DT.Recordset.RecordCount To DT.Recordset.RecordCount
        DT.Recordset.AddNew
        DT.Recordset!Nomor = i + 1
        DT.Recordset.Update
    Next i
End Function

Private Sub DataGrid1_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
End Sub

Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
    If DataGrid1.Col = 1 Then
        Call BukaDB
        RSBarang.Open "Select * from Barang where Kodebrg='" & DT.Recordset!Kode & "'", Conn
        If RSBarang.EOF Then
            Pesan = MsgBox("Kode Barang Tidak Terdaftar")
            List1.SetFocus
            Exit Sub
        End If
        DT.Recordset!Kode = RSBarang!KodeBrg
        DT.Recordset!Nama = RSBarang!NamaBrg
        DT.Recordset!Harga = RSBarang!HargaJual
        DataGrid1.Col = 4
        Exit Sub
    End If
    
    If DataGrid1.Col = 4 Then
        DT.Recordset!Jumlah = DT.Recordset!Jumlah
        DT.Recordset!Total = DT.Recordset!Harga * DT.Recordset!Jumlah
        DT.Recordset.Update
        Call Tambah_Baris
        DT.Recordset.MoveNext
        DataGrid1.Col = 1
        DT.Recordset.MoveLast
        DataGrid1.Refresh
        Total = TotalHarga
        JmlItem = TotalItem
    End If
End Sub

Function TotalHarga()
    Set TTlHarga = New ADODB.Recordset
    TTlHarga.Open "select sum(Total) as JumTotal from Transaksi", Conn
    TotalHarga = TTlHarga!JumTotal
End Function

Function TotalItem()
    Set TTlItem = New ADODB.Recordset
    TTlItem.Open "select sum(Jumlah) as JumItem from Transaksi", Conn
    TotalItem = TTlItem!Jumitem
End Function

Private Sub Bersihkan()
    JmlItem = ""
    Total = ""
    DP = ""
    Sisa = ""
    Stok = ""
End Sub

Sub Kosongksm()
NamaKsm = ""
AlamatKsm = ""
TeleponKsm = ""
End Sub

Private Sub DP_KeyPress(Keyascii As Integer)
    If Keyascii = 13 Then
        If DP = "" Then
            DP = 0
            Sisa = Total
        ElseIf DP = Total Then
            Sisa = 0
        ElseIf DP > Val(Total) Then
            MsgBox "Kembali : " & DP - Total & ""
            Sisa = 0
        ElseIf DP < Val(Total) Then
            Sisa = Total - DP
        End If
        
        CmdSimpan.Enabled = True
        CmdSimpan.SetFocus
    End If
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub CmdSimpan_Keypress(Keyascii As Integer)
    If Keyascii = 27 Then
        CmdSimpan.Enabled = False
        DP = ""
        DP.SetFocus
    End If
End Sub

Sub SimpanKsm()
Call BukaDB
RSKonsumen.Open "select * from konsumen where nomorksm='" & NomorKsm & "'", Conn
RSKonsumen.Requery
If RSKonsumen.EOF Then
    Dim SQLTambahksm As String
    SQLTambahksm = "Insert Into Konsumen(NomorKsm,namaksm,AlamatKsm,Teleponksm)" & _
    "values('" & NomorKsm & "','" & NamaKsm & "','" & AlamatKsm & "','" & TeleponKsm & "')"
    Conn.Execute (SQLTambahksm)
End If
End Sub

Private Sub CmdSimpan_Click()
       
    If NamaKsm = "" Or AlamatKsm = "" Or TeleponKsm = "" Then
        MsgBox "data pemesan belum lengkap"
        Exit Sub
    End If
    
    Dim Input1 As String
    Input1 = "Insert Into Pesanan(NomorPsn,TanggalPsn,Totalitem,TotalHrg,DP,Sisa,Nomorksm,Kodeksr,TglMintakrm,Ket)" & _
    "values('" & NomorPsn & "','" & Tanggal & "','" & JmlItem & "','" & Total & "','" & DP & "','" & Sisa & "','" & NomorKsm & "','" & Kodeksr & "','" & TglMintakrm & "','BELUM DIKIRIM')"
    Conn.Execute (Input1)
       
    RSTransaksi.Open "select * from Transaksi", Conn
    RSTransaksi.MoveFirst
    Do While Not RSTransaksi.EOF
        If RSTransaksi!Kode <> vbNullString Then
            Dim SQLTambahDetail As String
            SQLTambahDetail = "Insert Into DetailPsn(Nomorpsn,KodeBrg,Jumlahpsn) " & _
            "values ('" & NomorPsn & "','" & RSTransaksi!Kode & "','" & RSTransaksi!Jumlah & "')"
            Conn.Execute (SQLTambahDetail)
        End If
    RSTransaksi.MoveNext
    Loop
    
    Call SimpanKsm
    
    DT.Recordset.MoveFirst
    Do While Not DT.Recordset.EOF
        If DT.Recordset!Kode <> vbNullString Then
            Call BukaDB
            RSBarang.Open "Select * from Barang where Kodebrg='" & DT.Recordset!Kode & "'", Conn
            If Not RSBarang.EOF Then
                Dim Kurangi As String
                Kurangi = "update barang set jumlahbrg='" & RSBarang!JumlahBrg - DT.Recordset!Jumlah & "' where kodebrg='" & DT.Recordset!Kode & "'"
                Conn.Execute (Kurangi)
            End If
        End If
    DT.Recordset.MoveNext
    Loop
    Bersihkan
    Kosongksm
    Combo1.SetFocus
    form_Activate
    Call Cetak
End Sub

Private Sub CmdBatal_Click()
    Bersihkan
    Combo1.SetFocus
    form_Activate
End Sub

Private Sub CmdTutup_Click()
    Unload Me
End Sub

Function Cetak()
Call BukaDB
RSPesanan.Open "select * from Pesanan Where NomorPsn In(Select Max(NomorPsn)From Pesanan)Order By NomorPsn Desc", Conn
Tampilkan.Show
Dim JmlHarga, JmlJual, JmlHasil As Double
Dim MGrs As String
Tampilkan.Font = "Courier New"
Tampilkan.Print
Tampilkan.Print
RSKasir.Open "select * From Kasir where KodeKsr= '" & RSPesanan!Kodeksr & "'", Conn
RSKonsumen.Open "select * From Konsumen where Nomorksm= '" & RSPesanan!NomorKsm & "'", Conn
Tampilkan.Print Tab(5); "Nomor      :   "; RSPesanan!NomorPsn
Tampilkan.Print Tab(5); "Tanggal    :   "; Format(RSPesanan!tanggalpsn, "DD-MMMM-YYYY")
Tampilkan.Print Tab(5); "Kasir      :   "; RSKasir!Namaksr
MGrs = String$(33, "-")

Tampilkan.Print Tab(5); "Pemesan    :   "; RSKonsumen!NamaKsm
Tampilkan.Print Tab(5); "Alamat     :   "; RSKonsumen!AlamatKsm
Tampilkan.Print Tab(5); "Telepon    :   "; RSKonsumen!TeleponKsm

Tampilkan.Print Tab(5); MGrs
RSDetailPsn.Open "select * from detailpsn Where NomorPsn='" & RSPesanan!NomorPsn & "'", Conn
RSDetailPsn.MoveFirst
no = 0
Do While Not RSDetailPsn.EOF
    no = no + 1
    Set RSBarang = New ADODB.Recordset
    RSBarang.Open "select * From Barang where Kodebrg= '" & RSDetailPsn!KodeBrg & "'", Conn
    RSBarang.Requery
    Harga = RSBarang!HargaJual
    Jumlah = RSDetailPsn!JumlahPsn
    Hasil = Harga * Jumlah
    Tampilkan.Print Tab(5); no; Space(2); RSBarang!NamaBrg
    Tampilkan.Print Tab(10); RKanan(Jumlah, "##"); Space(1); "X";
    Tampilkan.Print Tab(15); Format(Harga, "###,###,###");
    Tampilkan.Print Tab(25); RKanan(Hasil, "###,###,###")
    RSDetailPsn.MoveNext
Loop
Tampilkan.Print Tab(5); MGrs
Tampilkan.Print Tab(5); "Total      :";
Tampilkan.Print Tab(25); RKanan(RSPesanan!TotalHrg, "###,###,###");
Tampilkan.Print Tab(5); "Uang Muka  :";
Tampilkan.Print Tab(25); RKanan(RSPesanan!DP, "###,###,###");

Tampilkan.Print Tab(5); MGrs
Tampilkan.Print Tab(5); "Sisa       :";
Tampilkan.Print Tab(25); RKanan(RSPesanan!Sisa, "###,###,###");

Tampilkan.Print Tab(5); MGrs
Tampilkan.Print
Tampilkan.Print
Tampilkan.Print
Conn.Close
End Function

Private Function RKanan(NData, CFormat) As String
    RKanan = Format(NData, CFormat)
    RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function

Private Sub List1_keyPress(Keyascii As Integer)
    If Keyascii = 13 Then
        If DataGrid1.SelText <> Left(List1, 6) Then
            DataGrid1.SelText = Left(List1, 6)
            DT.Recordset.Update
            Call BukaDB
            RSBarang.Open "Select * from Barang where KodeBrg='" & Left(List1, 6) & "'", Conn
            RSBarang.Requery
            If Not RSBarang.EOF Then
                DT.Recordset!Kode = RSBarang!KodeBrg
                DT.Recordset!Nama = RSBarang!NamaBrg
                DT.Recordset!Harga = RSBarang!HargaJual
                Stok = RSBarang!JumlahBrg
                DT.Recordset.Update
                DataGrid1.SetFocus
                DataGrid1.Col = 4
            End If
        End If
    End If
End Sub

