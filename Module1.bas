Attribute VB_Name = "Module1"
Public crx As New CRAXDRT.Application
Public conn As New ADODB.Connection
Public conpur As New ADODB.Connection
Public rsso As New ADODB.Recordset
Public rscode As New ADODB.Recordset
Public vread As New ADODB.Recordset
Public rschkcode As New ADODB.Recordset
Public rscompletion_slip As New ADODB.Recordset
Public rsdata_mesin As New ADODB.Recordset
Public strnamaserver As String

Public Sub db()
Set conn = New ADODB.Connection
Set rscode = New ADODB.Recordset
Set vread = New ADODB.Recordset
Set rscompletion_slip = New ADODB.Recordset
Set rsdata_mesin = New ADODB.Recordset

strnamadatabase = "purchasing"
strnamaserver = "HDGNGIT002\SQLEXPRESS"
strnamapemakai = "sa"
strpassword = "admin123"

'conn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;Initial Catalog=PURCHASING;Data Source =192.168.10.250, 1433;User Id=sa;Password=admin123"
conn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;Initial Catalog=PURCHASING;Data Source =" & strnamaserver & ";User Id=sa;Password=admin123"
'conn.ConnectionString = "Provider=" & strprovider & ";Data Source=" & strnamaserver & ";Initial Catalog=" & strnamadatabase & ";User ID=" & strnamapemakai & ";Password=" & strpassword
'conn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;Initial Catalog=penjadwalan_produksi;Data Source =192.168.0.108, 1433;User Id=sa;Password=admin123"
'conn.ConnectionString = strconstr
conn.CursorLocation = adUseClient
conn.Open
End Sub

Public Sub db_purchasing()
Set conn = New ADODB.Connection
Set rsso = New ADODB.Recordset
'conpur.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;Initial Catalog=purchasing;Data Source =192.168.10.250, 1433;User Id=sa;Password=admin123"
'connpur.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;Initial Catalog=penjadwalan_produksi;Data Source =192.168.0.108, 1433;User Id=sa;Password=admin123"
conpur.CursorLocation = adUseClient
conpur.Open
End Sub
