VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form DataPsn 
   Caption         =   "Rincian Pemesanan Barang"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8040
   BeginProperty Font 
      Name            =   "Century"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   4065
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   350
      Left            =   4920
      TabIndex        =   14
      Top             =   3600
      Width           =   2800
   End
   Begin VB.ListBox List1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   350
      Left            =   960
      TabIndex        =   7
      Top             =   2880
      Width           =   3000
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   350
      Left            =   960
      TabIndex        =   6
      Top             =   3240
      Width           =   3000
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   350
      Left            =   960
      TabIndex        =   5
      Top             =   3600
      Width           =   3000
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   350
      Left            =   4920
      TabIndex        =   4
      Top             =   2880
      Width           =   2800
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   350
      Left            =   4920
      TabIndex        =   3
      Top             =   3240
      Width           =   2800
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   350
      Left            =   6240
      TabIndex        =   2
      Top             =   2400
      Width           =   500
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   350
      Left            =   6720
      TabIndex        =   1
      Top             =   2400
      Width           =   1000
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1920
      Top             =   2400
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Caption         =   "Adodc1"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "DataPsn.frx":0000
      Height          =   2175
      Left            =   1920
      TabIndex        =   8
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3836
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Nama Barang"
         Caption         =   "Nama Barang"
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
      BeginProperty Column02 
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
      BeginProperty Column03 
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
            ColumnWidth     =   3000,189
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   615,118
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   915,024
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Ket"
      Height          =   345
      Left            =   4080
      TabIndex        =   15
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama"
      Height          =   345
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Alamat"
      Height          =   345
      Left            =   120
      TabIndex        =   12
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Telepon"
      Height          =   345
      Left            =   120
      TabIndex        =   11
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kasir"
      Height          =   345
      Left            =   4080
      TabIndex        =   10
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tgl Kirim"
      Height          =   345
      Left            =   4080
      TabIndex        =   9
      Top             =   3240
      Width           =   855
   End
End
Attribute VB_Name = "DataPsn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error Resume Next
'buka database
Call BukaDB
'bersihkan dulu list
List1.Clear
'cari nomor NomorPsn di tabel pesanan
RSPesanan.Open "Select Distinct NomorPsn from pesanan ", Conn
'tampilkan di list
Do Until RSPesanan.EOF
    List1.AddItem RSPesanan!NomorPsn
    RSPesanan.MoveNext
Loop
Conn.Close
End Sub

'ketika salah satu NomorPsn dipilih, maka...
Private Sub list1_click()
'buka database
Call BukaDB
Conn.CursorLocation = adUseClient
'cari data pesanan yang NomorPsnnya dipilih
RSPesanan.Open "select * from pesanan where NomorPsn='" & List1.Text & "'", Conn
RSPesanan.Requery
'jika ditemukan tampilkan TanggalPsnnya
If Not RSPesanan.EOF Then Text8 = RSPesanan!tanggalpsn: Text5 = RSPesanan!Ket
'buka tabel Konsumen yang ada di tabel pesanan sesuai noor NomorPsn
RSKonsumen.Open "select * from Konsumen where NomorKsm='" & RSPesanan!NomorKsm & "'", Conn
'jika ditemukan tampilkan data-datanya
If Not RSKonsumen.EOF Then
    Text2 = RSKonsumen!NamaKsm
    Text3 = RSKonsumen!AlamatKsm
    Text4 = RSKonsumen!TeleponKsm
End If
'buka tabel kasir yang kodenya disimpan di tabel pesanan berdasarkan nomor NomorPsn
RSKasir.Open "select * from Kasir where KodeKsr='" & RSPesanan!Kodeksr & "'", Conn
'jika ditemukan tampilkan kode dan nama kasir
If Not RSKasir.EOF Then
    Text7 = RSKasir!Namaksr
End If

Conn.Close
'hubungkan objek adodc ke database
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOPesanan.mdb"
'tampilkan nama barang, harga Pesanan, jumlah Pesanan dan total di tabel pesanan,detail Pesanan yang NomorPsnnya dipilih dalam list
Adodc1.RecordSource = "select NamaBrg as [Nama Barang], HargaJual as Harga,JumlahPsn as Jumlah, HargaJual*JumlahPsn as Total from Barang,DetailPsn where DetailPsn.kodeBrg=Barang.kodeBrg and NomorPsn='" & List1.Text & "'"
Adodc1.Refresh
'hubungkan datagrid1 dengan objek adodc
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
'tampilkan total dan item
Call Total
Call Item
End Sub

Private Sub List1_keyPress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
End Sub

'mencari jumlah total item barang
Function Item()
Adodc1.Recordset.MoveFirst
Jumlah = 0
Do While Not Adodc1.Recordset.EOF
    Jumlah = Jumlah + Adodc1.Recordset!Jumlah
    Adodc1.Recordset.MoveNext
Loop
Text9 = Jumlah
End Function

'mencari jumlah total harga Pesanan
Function Total()
Adodc1.Recordset.MoveFirst
Jumlah = 0
Do While Not Adodc1.Recordset.EOF
    Jumlah = Jumlah + Adodc1.Recordset!Total
    Adodc1.Recordset.MoveNext
Loop
Text10 = Jumlah
End Function

