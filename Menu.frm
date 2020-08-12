VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Menu 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Menu Utama"
   ClientHeight    =   3390
   ClientLeft      =   195
   ClientTop       =   765
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
   Picture         =   "Menu.frx":0000
   ScaleHeight     =   3390
   ScaleWidth      =   4275
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   2895
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   873
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Crystal.CrystalReport CR 
      Left            =   1440
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1800
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   13
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":42DB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":430CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":433E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":436FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":43A19
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":43D33
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":4404D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":44367
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":44681
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":4499B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":44CB5
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":44FCF
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":452E9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnfile 
      Caption         =   "File"
      Begin VB.Menu mnkasir 
         Caption         =   "Kasir"
      End
      Begin VB.Menu mnbarang 
         Caption         =   "Barang"
      End
      Begin VB.Menu mnkurir 
         Caption         =   "Kurir"
      End
   End
   Begin VB.Menu mntransaksi 
      Caption         =   "Transaksi"
      Begin VB.Menu mnpesanan 
         Caption         =   "Pesanan"
      End
      Begin VB.Menu mnpengiriman 
         Caption         =   "Pengiriman"
      End
   End
   Begin VB.Menu mnlaporan 
      Caption         =   "Laporan"
      Begin VB.Menu mnlapbarang 
         Caption         =   "Data Barang"
      End
      Begin VB.Menu mnlappesanan 
         Caption         =   "Rincian Pemesanan"
      End
      Begin VB.Menu mnlappengiriman 
         Caption         =   "Rincian Pengiriman"
      End
      Begin VB.Menu mnakumpsn 
         Caption         =   "Akumulasi Pemesanan"
      End
      Begin VB.Menu mnkumkrm 
         Caption         =   "Akumulasi Pengiriman"
      End
   End
   Begin VB.Menu mnjejak 
      Caption         =   "Jejak Transaksi"
      Begin VB.Menu mnjejakpsn 
         Caption         =   "Pemesanan"
      End
      Begin VB.Menu mnjejakkrm 
         Caption         =   "Pengiriman"
      End
   End
   Begin VB.Menu mnkeluar 
      Caption         =   "Keluar"
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub mnakumpsn_Click()
AkumPsn.Show
End Sub

Private Sub mnbarang_Click()
Barang.Show
End Sub

Private Sub mnjejakkrm_Click()
DataKrm.Show
End Sub

Private Sub mnjejakpsn_Click()
DataPsn.Show
End Sub

Private Sub mnkasir_Click()
Kasir.Show
End Sub

Private Sub mnkeluar_Click()
End
End Sub

Private Sub mnkumkrm_Click()
AkumKrm.Show
End Sub

Private Sub mnkurir_Click()
Kurir.Show
End Sub

Private Sub mnlapbarang_Click()
CR.ReportFileName = App.Path & "\Lap Barang.rpt"
CR.WindowState = crptMaximized
CR.RetrieveDataFiles
CR.Action = 1
End Sub

Private Sub mnlappengiriman_Click()
LapKirim.Show
End Sub

Private Sub mnlappesanan_Click()
LapPesan.Show
End Sub

Private Sub mnpengiriman_Click()
Pengiriman.Show
End Sub

Private Sub mnpesanan_Click()
Pesanan.Show
End Sub

Private Sub mnujisql_Click()
UjiSQL.Show
End Sub

Private Sub nlaporan_Click()

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Key
    Case "a"
        Kasir.Show
    Case "b"
        Barang.Show
    Case "c"
        Kurir.Show
    Case "d"
        Pesanan.Show
    Case "e"
        Pengiriman.Show
    Case "f"
       CR.ReportFileName = App.Path & "\Lap Barang.rpt"
        CR.WindowState = crptMaximized
        CR.RetrieveDataFiles
        CR.Action = 1
    Case "g"
        LapPesan.Show
    Case "h"
        LapKirim.Show
    Case "i"
        AkumPsn.Show
    Case "j"
        AkumPsn.Show
    Case "k"
        DataPsn.Show
    Case "l"
        DataKrm.Show
    Case "m"
        End
End Select
End Sub
