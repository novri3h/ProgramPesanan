VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Pengiriman 
   Caption         =   "Data Pengiriman"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7725
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
   ScaleHeight     =   5835
   ScaleWidth      =   7725
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      Height          =   345
      Left            =   5880
      TabIndex        =   1
      Top             =   1800
      Width           =   1500
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Pengiriman.frx":0000
      Height          =   1815
      Left            =   120
      TabIndex        =   29
      Top             =   2640
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   3201
      _Version        =   393216
      AllowUpdate     =   0   'False
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
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Dibayar 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   6600
      TabIndex        =   2
      Top             =   5040
      Width           =   1000
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   1500
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   350
      Left            =   120
      TabIndex        =   3
      Top             =   4680
      Width           =   850
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   350
      Left            =   960
      TabIndex        =   4
      Top             =   4680
      Width           =   850
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   350
      Left            =   1800
      TabIndex        =   5
      Top             =   4680
      Width           =   850
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2160
      Top             =   5160
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identitas Pemesan"
      Height          =   1935
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   4335
      Begin VB.Label TeleponKsm 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1200
         TabIndex        =   39
         Top             =   720
         Width           =   3000
      End
      Begin VB.Label AlamatKsm 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1200
         TabIndex        =   38
         Top             =   1440
         Width           =   3000
      End
      Begin VB.Label NamaKsm 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1200
         TabIndex        =   37
         Top             =   1080
         Width           =   3000
      End
      Begin VB.Label NomorKsm 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1200
         TabIndex        =   36
         Top             =   360
         Width           =   3000
      End
      Begin VB.Label Label17 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nomor"
         Height          =   300
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Label18 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nama"
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label Label19 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Alamat"
         Height          =   300
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label Label20 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Telepon"
         Height          =   300
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1005
      End
   End
   Begin MSAdodcLib.Adodc DT 
      Height          =   405
      Left            =   120
      Top             =   5160
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
   Begin VB.Label NamaKrr 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   5880
      TabIndex        =   35
      Top             =   2160
      Width           =   1500
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama Kurir"
      Height          =   300
      Left            =   4800
      TabIndex        =   34
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Kembali 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   6600
      TabIndex        =   33
      Top             =   5400
      Width           =   1005
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kembali"
      Height          =   300
      Left            =   5520
      TabIndex        =   32
      Top             =   5400
      Width           =   1005
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
      Left            =   4800
      TabIndex        =   31
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label TglMintakrm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   5880
      TabIndex        =   30
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode Kurir"
      Height          =   300
      Left            =   4800
      TabIndex        =   28
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label DP 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   4440
      TabIndex        =   27
      Top             =   5040
      Width           =   1005
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Dibayar"
      Height          =   300
      Left            =   5520
      TabIndex        =   26
      Top             =   5040
      Width           =   1005
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nomor Psn"
      Height          =   315
      Left            =   240
      TabIndex        =   25
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nomor Krm"
      Height          =   300
      Left            =   4800
      TabIndex        =   24
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kasir"
      Height          =   300
      Left            =   4800
      TabIndex        =   23
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal"
      Height          =   300
      Left            =   4800
      TabIndex        =   22
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Kodeksr 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3600
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Nomorkrm 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   5880
      TabIndex        =   20
      Top             =   720
      Width           =   1500
   End
   Begin VB.Label Namaksr 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   5880
      TabIndex        =   19
      Top             =   1440
      Width           =   1500
   End
   Begin VB.Label Tanggal 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   5880
      TabIndex        =   18
      Top             =   1080
      Width           =   1500
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total"
      Height          =   300
      Left            =   3360
      TabIndex        =   17
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
      Left            =   3360
      TabIndex        =   16
      Top             =   5040
      Width           =   1005
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Sisa"
      Height          =   300
      Left            =   5520
      TabIndex        =   15
      Top             =   4680
      Width           =   1005
   End
   Begin VB.Label Total 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   4440
      TabIndex        =   14
      Top             =   4680
      Width           =   1005
   End
   Begin VB.Label Sisa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   6600
      TabIndex        =   13
      Top             =   4680
      Width           =   1005
   End
   Begin VB.Label JmlItem 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   2760
      TabIndex        =   12
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Item"
      Height          =   300
      Left            =   2760
      TabIndex        =   11
      Top             =   4680
      Width           =   500
   End
End
Attribute VB_Name = "Pengiriman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub form_Activate()
Call BukaDB
DT.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOPesanan.mdb"
DT.RecordSource = "Transaksi1"
Set DataGrid1.DataSource = DT
DataGrid1.Refresh
    
If Kodeksr = "" Then
    MsgBox "Kasir tidak terdeteksi"
    Login.Show
    Exit Sub
End If

Call Autokrm
Call Tabel_Kosong
Tanggal = Format(Date, "dd-mm-yyyy")
CmdSimpan.Enabled = False
Combo1.SetFocus

Call BukaDB
RSPesanan.Open "Select * from Pesanan where ket='BELUM DIKIRIM'", Conn
Combo1.Clear
Do Until RSPesanan.EOF
    Combo1.AddItem RSPesanan!NomorPsn
    RSPesanan.MoveNext
Loop
    
RSKurir.Open "Select * from kurir ", Conn
Combo2.Clear
Do Until RSKurir.EOF
    Combo2.AddItem RSKurir!KodeKrr
    RSKurir.MoveNext
Loop
Conn.Close
End Sub

Private Sub Form_Load()
    Kodeksr = Login.TxtKodeKsr
    Namaksr = Login.TxtNamaKsr
    DataGrid1.Col = 1
    CmdSimpan.Enabled = False
End Sub

Private Sub Autokrm()
Call BukaDB
RSPengiriman.Open ("select * from pengiriman Where Nomorkrm In(Select Max(Nomorkrm)From Pengiriman)Order By Nomorkrm Desc"), Conn
RSPengiriman.Requery
    Dim Urutan As String * 10
    Dim Hitung As Long
    With RSPengiriman
        If .EOF Then
            Urutan = "K" + Format(Date, "yymmdd") + "001"
            Nomorkrm = Urutan
        Else
            If Mid(!Nomorkrm, 2, 6) <> Format(Date, "yymmdd") Then
                Urutan = "K" + Format(Date, "yymmdd") + "001"
            Else
                Hitung = Mid(!Nomorkrm, 9) + 1
                Urutan = "K" + Format(Date, "yymmdd") + Right("000" & Hitung, 3)
            End If
        End If
        Nomorkrm = Urutan
    End With
End Sub

Function Tabel_Kosong()
If DT.Recordset.RecordCount > 0 Then
    DT.Recordset.MoveFirst
    Do While Not DT.Recordset.EOF
        DT.Recordset.Delete
        DT.Recordset.MoveNext
    Loop
End If
End Function

Private Sub Combo1_Keypress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
If Keyascii = 13 Then
    Call BukaDB
    RSPesanan.Open "Select * from Pesanan where nomorpsn='" & Combo1 & "'", Conn
    RSPesanan.Requery
    If RSPesanan.EOF Then
        MsgBox "Nomor pesanan tidak terdaftar"
        Combo1.SetFocus
        Exit Sub
    Else
        Combo2.SetFocus
    End If
End If

End Sub

Private Sub Combo1_Click()
Call BukaDB
RSPesanan.Open "Select * from Pesanan where nomorpsn='" & Combo1 & "'", Conn
RSPesanan.Requery
If Not RSPesanan.EOF Then
    TglMintakrm = CDate(RSPesanan!TglMintakrm)
    Total = Format(RSPesanan!TotalHrg, "###,###,###")
    Sisa = Format(RSPesanan!Sisa, "###,###,###")
    DP = Format(RSPesanan!DP, "###,###,###")
    
    If Total = Sisa Then
        DP = 0
    Else
        DP = Format(RSPesanan!DP, "###,###,###")
    End If
    
    If Val(DP) >= Val(Total) Then
        Sisa = 0
        Kembali = 0
    Else
        Sisa = Format(RSPesanan!Sisa, "###,###,###")
    End If
    
    JmlItem = Val(RSPesanan!TotalItem)
    NomorKsm = RSPesanan!NomorKsm
    Dim RS As New ADODB.Recordset
    RS.Open "select barang.kodebrg,barang.namabrg,barang.hargajual,jumlahpsn,hargajual*jumlahpsn as total  from barang,detailpsn where left(nomorpsn,10)='" & Combo1 & "' and barang.kodebrg=detailpsn.kodebrg", Conn
    Call Tabel_Kosong
    RS.MoveFirst
    Nomor = 0
    Do While Not RS.EOF
        Nomor = Nomor + 1
        DT.Recordset.AddNew
        DT.Recordset!Nomor = Nomor
        DT.Recordset!Kode = RS!KodeBrg
        DT.Recordset!Nama = RS!NamaBrg
        DT.Recordset!Harga = RS!HargaJual
        DT.Recordset!Jumlah = RS!JumlahPsn
        DT.Recordset!Total = RS!Total
        DT.Recordset.Update
        RS.MoveNext
    Loop
Else
    MsgBox "nomor pesanan tidak terdaftar"
    Combo1.SetFocus
    Exit Sub
End If
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

Private Sub Combo2_click()
Call BukaDB
RSKurir.Open "select * from kurir where kodekrr='" & Combo2 & "'", Conn
If Not RSKurir.EOF Then
    NamaKrr = RSKurir!NamaKrr
Else
    MsgBox "kode kurir tidak terdaftar"
    Combo2.SetFocus
End If
Conn.Close
End Sub

Private Sub Combo2_Keypress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    Call BukaDB
    RSKurir.Open "select * from kurir where kodekrr='" & Combo2 & "'", Conn
    If Not RSKurir.EOF Then
        NamaKrr = RSKurir!NamaKrr
    Else
        MsgBox "kode kurir tidak terdaftar"
        Combo2.SetFocus
        Exit Sub
    End If
    
    If Val(DP) >= Val(Total) Then
        Dibayar.Enabled = False
        Dibayar = 0
        CmdSimpan.Enabled = True
        CmdSimpan.SetFocus
    Else
        Dibayar.Enabled = True
        Dibayar.SetFocus
    End If
End If
End Sub

Private Sub Dibayar_KeyPress(Keyascii As Integer)
    If Keyascii = 13 Then
        If Dibayar = "" Or Val(Dibayar) < (Sisa) Then
            MsgBox "Jumlah Pembayaran Kurang"
            Dibayar.SetFocus
        Else
            Dibayar = Format(Dibayar, "###,###,###")
            If Dibayar = Sisa Then
                Kembali = Dibayar - Sisa
            Else
                Kembali = Format(Dibayar - Sisa, "###,###,###")
            End If
        CmdSimpan.Enabled = True
        CmdSimpan.SetFocus
        End If
    End If
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub Timer1_Timer()
    Jam = Time$
End Sub

Private Sub Bersihkan()
    Combo1 = ""
    JmlItem = ""
    Total = ""
    DP = ""
    Sisa = ""
    TglMintakrm = ""
    Combo2 = ""
    Dibayar = ""
    NomorKsm = ""
    NamaKsm = ""
    AlamatKsm = ""
    TeleponKsm = ""
    Kembali = ""
    NamaKrr = ""
End Sub

Private Sub CmdSimpan_Click()
    If Combo1 = "" Or Combo2 = "" Then
        MsgBox "data pengiriman belum lengkap"
        Exit Sub
    Else
        If Sisa <> 0 And Dibayar = "" Then
            MsgBox "Pembayaran belum lunas"
            Dibayar.SetFocus
            Exit Sub
        End If
    End If
        
    Call BukaDB
    'simpan ke tabel pengiriman
    Dim SimpanPesanan As String
    SimpanPesanan = "Insert Into Pengiriman(Nomorkrm,Nomorpsn,Tanggalkrm,Total,DP,Sisa,Dibayar,Kembali,Nomorksm,Kodeksr,Kodekrr)" & _
    "values('" & Nomorkrm & "','" & Combo1 & "','" & Tanggal & "','" & Total & "','" & DP & "','" & Sisa & "','" & Dibayar & "','" & Kembali & "','" & NomorKsm & "','" & Kodeksr & "','" & Combo2 & "')"
    Conn.Execute (SimpanPesanan)
    
    'ubah ket di tabel pesanan
    Dim SimpanPesanan1 As String
    SimpanPesanan1 = "Update Pesanan set Ket='TELAH DIKIRIM' where nomorpsn='" & Combo1 & "'"
    Conn.Execute (SimpanPesanan1)
    
    'simpan ke tabel detailkrm
    DT.Recordset.MoveFirst
    Do While Not DT.Recordset.EOF
        Dim SimpanDetailPsn As String
        SimpanDetailPsn = "Insert Into Detailkrm(Nomorkrm,KodeBrg,Jumlahkrm) " & _
        "values ('" & Nomorkrm & "','" & DT.Recordset!Kode & "','" & DT.Recordset!Jumlah & "')"
        Conn.Execute (SimpanDetailPsn)
    DT.Recordset.MoveNext
    Loop
       
    Bersihkan
    form_Activate
End Sub

Private Sub CmdBatal_Click()
'Conn.Close
Bersihkan
form_Activate
End Sub

Private Sub CmdTutup_Click()
    Unload Me
End Sub

