Attribute VB_Name = "Module1"
Public conn As New ADODB.Connection
Public noidsepatu As ADODB.Recordset
Public noidpembayaran As ADODB.Recordset


Public Sub buka()
    Set conn = New ADODB.Connection
    Set noidsepatu = New ADODB.Recordset
    Set noidpembayaran = New ADODB.Recordset
    
    conn.Open "provider=microsoft.jet.oledb.4.0;data source =" & App.Path & "\db_penjualan.mdb"
    
End Sub
