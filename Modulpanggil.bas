Attribute VB_Name = "Module1"
Public cn As New ADODB.Connection
Public Sub bukakoneksi()
strkoneksi = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
            "Data Source=" & App.Path & "\data\NILAI.mdb;" & _
            "Persist Security Info=False "
cn.Open strkoneksi
If cn.State = 1 Then
    MsgBox " Koneksi Sukses", vbInformation
Else
    MsgBox " Koneksi Gagal", vbInformation
End If
End Sub



