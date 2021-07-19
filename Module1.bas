Attribute VB_Name = "Module1"
Option Explicit
'variables para conexion a la base de datos
Global CN As New ADODB.Connection
Public a As Double
Public s As Double
Global q As String
Global RSVENTAS_ELIMINADAS As New ADODB.Recordset
Global RSFACTURA_ELIMINADAS As New ADODB.Recordset
Global RSINV As New ADODB.Recordset
Global RSVEN As New ADODB.Recordset
Global rsFactura As New ADODB.Recordset
'variable para acceder a la tabla PROVEEDORES
Global RSPRO As New ADODB.Recordset
Global RSNOM As New ADODB.Recordset
Global RSPROD As New ADODB.Recordset
Global privilegiosadmin As Integer


'procedimiento principal
'conexion a la base de datos
Sub MAIN()
    With CN
        .CursorLocation = adUseClient 'Vamos a ser clientes de la base de datos
        'Conexion a la base de datos
        .Open "Provider=Microsoft.Jet.OLEDB.4.0;" & " Data Source= " & App.Path & "\DATA\BASEINV.mdb;Persist Security Info=False"
        'frmDetallesLibro.Show
        FRMVENTAS.Show
        
    End With
End Sub
Sub TABLAPRODUCTO()
With RSPROD
    If .State = 1 Then .Close
    .Source = "INVENTARIO"
         .CursorType = adOpenKeyset 'Definimos el tipo de cursor.
        .LockType = adLockBatchOptimistic 'Definimos el tipo de bloqueo.
        .Open "select* from INVENTARIO", CN
     End If
     RSPROD.MoveFirst
End Sub
Sub TABLANOMBRE()
With RSNOM
    If .State = 1 Then .Close
    .Source = "PROVEEDORES"
    .CursorType = adOpenKeyset 'Definimos el tipo de cursor.
        .LockType = adLockBatchOptimistic 'Definimos el tipo de bloqueo.
        .Open "select* from PROVEEDORES", CN
        RSNOM.MoveFirst
        End With
End Sub
Sub tablaPROVEEDORES()
    With RSPRO
    
        If .State = 1 Then .Close
        .Source = "PROVEEDORES"
        .CursorType = adOpenKeyset 'Definimos el tipo de cursor.
        .LockType = adLockBatchOptimistic 'Definimos el tipo de bloqueo.
        .Open "select* from PROVEEDORES", CN
        End With
    RSPRO.MoveFirst
    
        
End Sub
Sub tablaINVENTARIO()
    With RSINV
        
        If .State = 1 Then .Close
        .Source = "INVENTARIO"
        .CursorType = adOpenKeyset 'Definimos el tipo de cursor.
        .LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
        .Open "select * from INVENTARIO", CN
    End With
    
    RSINV.MoveFirst
    
End Sub
Sub tablaVENTAS()
    With RSVEN
        
        If .State = 1 Then .Close
        .Source = "VENTAS"
        .CursorType = adOpenKeyset 'Definimos el tipo de cursor.
        .LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
        .Open "select * from VENTAS", CN
    End With
    
    
End Sub
Sub factura()
With rsFactura
    If .State = 1 Then .Close
    .CursorType = adOpenKeyset 'Definimos el tipo de cursor.
     .LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
    .Open "select * from FACTURA", CN
    End With
End Sub

Sub VENTAS_ELIMINADAS()
    With RSVENTAS_ELIMINADAS
        If .State = 1 Then .Close
        .Open "select * from VENTAS_ELIMINADAS", CN, adOpenStatic, adLockOptimistic
    End With
End Sub
    
Sub FACTURA_ELIMINADA()
     With RSFACTURA_ELIMINADAS
     If .State = 1 Then .Close
     .Open "select * from FACTURA_ELIMINADA", CN, adOpenStatic, adLockOptimistic
     End With
End Sub

