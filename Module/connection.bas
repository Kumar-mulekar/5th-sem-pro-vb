Attribute VB_Name = "Module2"
Public con As ADODB.Connection
Public rec As ADODB.Recordset

Public Sub main()
Set con = New ADODB.Connection

con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\HYDRA\Desktop\Sem-5-pro\Databases\ProData.mdb;Persist Security Info=False"
con.CursorLocation = adUseClient

End Sub

