Attribute VB_Name = "MdlPesanan"

Public Conn As New ADODB.Connection
Public RSBarang As ADODB.Recordset
Public RSKasir As ADODB.Recordset
Public RSKonsumen As ADODB.Recordset
Public RSPesanan As ADODB.Recordset
Public RSDetailPsn As ADODB.Recordset
Public RSKurir As ADODB.Recordset
Public RSPengiriman As ADODB.Recordset
Public RSDetailKrm As ADODB.Recordset
Public RSTransaksi As ADODB.Recordset

Public Sub BukaDB()
Dim STR As String
Set Conn = New ADODB.Connection
Set RSBarang = New ADODB.Recordset
Set RSKasir = New ADODB.Recordset
Set RSKonsumen = New ADODB.Recordset
Set RSPesanan = New ADODB.Recordset
Set RSDetailPsn = New ADODB.Recordset
Set RSKurir = New ADODB.Recordset
Set RSPengiriman = New ADODB.Recordset
Set RSDetailKrm = New ADODB.Recordset
Set RSTransaksi = New ADODB.Recordset
Conn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ADOPesanan.mdb"
End Sub


