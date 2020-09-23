Attribute VB_Name = "MdlInventory"
Public cn As ADODB.Connection
Public RSkredit As ADODB.Recordset
Public RSdeveloper As ADODB.Recordset
Public RSnotaris As ADODB.Recordset
Public RSnomer As ADODB.Recordset
Public RStambah As ADODB.Recordset
Dim koneksi As String
Sub Main()
frmsplash.Show
End Sub

Public Sub Bukakoneksi()
Set cn = New ADODB.Connection
      koneksi = "Provider=SQLOLEDB.1;" & _
                "Initial Catalog=DAJAM; " & _
                "User ID=dajam; Password=;" & _
                "Data Source=;"
    cn.Open koneksi
End Sub


