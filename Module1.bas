Attribute VB_Name = "Module1"
Global a As String
Global CN As New ADODB.Connection
Global RSINV As New ADODB.Recordset
Global RSVEN As New ADODB.Recordset
Sub main()
    With CN
        .CursorLocation = adUseClient 'Vamos a ser clientes de la base de datos
        'Conexion a la base de datos
        .Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BASEINV.mdb"
        'frmDetallesLibro.Show
        FRMLOGIN.Show
        
    End With
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
    
    RSVEN.MoveFirst
    
End Sub

