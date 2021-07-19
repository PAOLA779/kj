Attribute VB_Name = "Module2"
Global B As String
Global CN As New ADODB.Connection
Global RSPRO As New ADODB.Recordset

Sub MAIN()
With CN
.CursorLocation = adUseClient
'Conexion a la base de datos
End Sub
