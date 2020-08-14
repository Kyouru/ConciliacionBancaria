Private UltimaFunc As String
Public UltimaFecha As Date
Public idPendiente As Integer

Public Sub btActualizar_Click()
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    
    UltimaFecha = fechaStrStr(tbFecha.Text)
    [FECHA1] = UltimaFecha
    [DIA] = StrConv(Format(UltimaFecha, "DDDD"), vbProperCase)
    Dim idcuenta As Integer
    
    OpenDB
    idcuenta = idCuentaPorNumero(rs, cnn, [CUENTA])
    
    If idcuenta = 0 Then
        MsgBox "Error, cuenta no encontrada"
        Exit Sub
    End If
    
    'Saldo Banco
    strSQL = "SELECT SALDO_BANCO, FECHA " & _
                "FROM DB_SALDO_BANCO " & _
                    "WHERE ID_CUENTA_FK = " & idcuenta & _
                        " AND FECHA = (SELECT MAX(FECHA) " & _
                    "FROM DB_SALDO_BANCO " & _
                        "WHERE ID_CUENTA_FK = " & idcuenta & " AND FECHA <= #" & fechaDateStr(UltimaFecha) & "#)"
    
    
    On Error Resume Next
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    'Log del Error
    logError cnn, Me.Name, "btActualizar_Click(1)"
    
    If rs.RecordCount > 0 Then
        [SALDO_BANCO] = rs.Fields("SALDO_BANCO")
        [FECHA_SALDO_BANCO] = rs.Fields("FECHA")
    End If
    
    'Saldo Sistema
    strSQL = "SELECT SALDO_SISTEMA, FECHA " & _
                "FROM DB_SALDO_SISTEMA " & _
                    "WHERE ID_CUENTA_FK = " & idcuenta & _
                        " AND FECHA = (SELECT MAX(FECHA) " & _
                    "FROM DB_SALDO_SISTEMA " & _
                        "WHERE ID_CUENTA_FK = " & idcuenta & " AND FECHA <= #" & fechaDateStr(UltimaFecha) & "#)"
    
    Set rs = Nothing
    On Error Resume Next
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    'Log del Error
    logError cnn, Me.Name, "btActualizar_Click(1)"
    
    If rs.RecordCount > 0 Then
        [SALDO_SISTEMA] = rs.Fields("SALDO_SISTEMA")
        [FECHA_SALDO_SISTEMA] = rs.Fields("FECHA")
    End If
    
    
    
    'Saldo Banco Dia Anterior
    strSQL = "SELECT SALDO_BANCO " & _
                "FROM DB_SALDO_BANCO " & _
                    "WHERE ID_CUENTA_FK = " & idcuenta & _
                        " AND FECHA = (SELECT MAX(FECHA) " & _
                                        "FROM DB_SALDO_BANCO " & _
                                        "WHERE ID_CUENTA_FK = " & idcuenta & " AND FECHA < #" & fechaDateStr(UltimaFecha) & "#)"
    
    
    Set rs = Nothing
    On Error Resume Next
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    'Log del Error
    logError cnn, Me.Name, "btActualizar_Click(2)"
    
    If rs.RecordCount > 0 Then
        [SALDO_BANCO_ANTERIOR] = rs.Fields("SALDO_BANCO")
    End If
    
    
    'Pendientes
    strSQL = "SELECT ID_PENDIENTE, FECHA_INICIO, IMPORTE, TIPO, DETALLE, ADICIONAL FROM PENDIENTE WHERE ID_CUENTA_FK = " & idcuenta & " AND FECHA_INICIO <= #" & fechaDateStr(UltimaFecha) & "# AND (FECHA_FIN IS NULL OR FECHA_FIN > #" & fechaDateStr(UltimaFecha) & "#) AND ANULADO = FALSE ORDER BY FECHA_INICIO ASC, ID_PENDIENTE ASC"

    Set rs = Nothing
    On Error Resume Next
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    'Log del Error
    logError cnn, Me.Name, "btActualizar_Click(3)"
    
    Range(Range("dataSet"), Range("dataSet")(1, 3).End(xlDown)).ClearContents
    
    If rs.RecordCount > 0 Then
        Me.Range("dataSet").CopyFromRecordset rs
    End If
    
    
    'Movimientos
    strSQL = "SELECT ID_MOVIMIENTO, MONTO, GLOSA, NOMBRE_TIPO_MOVIMIENTO, NUMERO_OPERACION, FECHA_MOVIMIENTO, HORA_MOVIMIENTO FROM MOVIMIENTO LEFT JOIN TIPO_MOVIMIENTO ON TIPO_MOVIMIENTO.ID_TIPO_MOVIMIENTO = MOVIMIENTO.ID_TIPO_MOVIMIENTO_FK WHERE ID_CUENTA_FK = " & [ID_CUENTA] & " AND FECHA_MOVIMIENTO = #" & fechaDateStr(UltimaFecha) & "#"
    
    Set rs = Nothing
    On Error Resume Next
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    'Log del Error
    logError cnn, Me.Name, "btActualizar_Click(4)"
    
    Range(Range("dataSetMovimientos"), Range("dataSetMovimientos").End(xlDown)).ClearContents
    
    If rs.RecordCount > 0 Then
        Me.Range("dataSetMovimientos").CopyFromRecordset rs
    End If
    closeRS
    
    Me.Range("D:D").EntireColumn.NumberFormat = "#,##0.00"
    Me.Range("E:H").EntireColumn.NumberFormat = "@"
    Me.Range("K:K").EntireColumn.NumberFormat = "#,##0.00"
    Me.Range("L:N").EntireColumn.NumberFormat = "@"
    Me.Range("FECHA_SALDO_BANCO").NumberFormat = "DD/MM/YYYY"
    Me.Range("FECHA_SALDO_SISTEMA").NumberFormat = "DD/MM/YYYY"
    Me.Range("SALDO_BANCO_ANTERIOR").EntireColumn.NumberFormat = "#,##0.00"
    Me.Range("SALDO_CALCULADO").EntireColumn.NumberFormat = "#,##0.00"
    Me.Range("CALCULADO_DIFERENCIA").EntireColumn.NumberFormat = "#,##0.00"
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Public Sub btBuscarSuma_Click()
    Hoja10.Activate
End Sub

Public Sub btCalendario_Click()
    UltimaFunc = "btCalendario_Click"
    Dim strFecha As String
    frmCalendario.Show
    If Not IsNull(UltimaFecha) Then
        Hoja1.tbFecha.Text = Format(UltimaFecha, "DD/MM/YYYY")
    End If
    btActualizar_Click
End Sub

Private Sub btEditarSaldoBanco_Click()
    UltimaFunc = "btEditarSaldoBanco_Click"
    frmEditarSaldoBanco.Show
End Sub

'Buscar el ID de la cuenta
Private Function idCuentaPorNumero(recos As ADODB.Recordset, conn As ADODB.Connection, nroCuenta As String)
    idCuentaPorNumero = 0
    strSQL = "SELECT ID_CUENTA FROM CUENTA WHERE NUMERO_CUENTA_SISGO = '" & nroCuenta & "'"
    
    On Error Resume Next
    recos.Open strSQL, conn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    'Log del Error
    logError conn, Me.Name, "idCuentaPorNumero"
    
    If recos.RecordCount > 0 Then
        idCuentaPorNumero = recos.Fields(0)
    End If
    Set recos = Nothing
End Function

Public Sub btCompararListas_Click()
    Hoja11.Activate
End Sub

Public Sub btMantenimiento_Click()
    frmMantenimientoPendientes.Show
End Sub

Private Sub btSaldoDiario_Click()

    Hoja5.tbFechaSaldoDiario.Text = tbFecha.Text
    Hoja5.ActualizarSaldos
    Hoja5.Activate
    
End Sub

Public Sub btPendientes_Click()
    UltimaFecha = fechaStrStr(tbFecha.Text)
    strSQL = "SELECT ID_PENDIENTE, NOMBRE_BANCO, '', NUMERO_CUENTA_SISGO, LEFT(NOMBRE_MONEDA,1), FECHA_INICIO, IMPORTE, TIPO, DETALLE FROM ((PENDIENTE LEFT JOIN CUENTA ON CUENTA.ID_CUENTA = PENDIENTE.ID_CUENTA_FK) LEFT JOIN BANCO ON BANCO.ID_BANCO = CUENTA.ID_BANCO_FK) LEFT JOIN MONEDA ON MONEDA.ID_MONEDA = CUENTA.ID_MONEDA_FK WHERE FECHA_INICIO <= #" & fechaDateStr(UltimaFecha) & "# AND (FECHA_FIN IS NULL OR FECHA_FIN > #" & fechaDateStr(UltimaFecha) & "#) AND ANULADO = FALSE ORDER BY NOMBRE_BANCO, ID_CUENTA, FECHA_INICIO"
    
    OpenDB
    
    On Error Resume Next
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    'Log del Error
    logError cnn, Me.Name, "btPendientes_Click"
    
    Hoja5.Range(Hoja5.Range("dataSetPendientes"), Hoja5.Range("dataSetPendientes").End(xlDown)).ClearContents
        
    'On Error Resume Next
    If rs.RecordCount > 0 Then
        Hoja5.Range("PENDIENTES_FECHA") = "PENDIENTES AL " & tbFecha.Text
        Hoja5.Activate
        Hoja5.Range("dataSetPendientes").CopyFromRecordset rs
        Hoja5.Range("dataSetPendientes")(1, 6).EntireColumn.NumberFormat = "DD/MM/YYYY"
        Hoja5.Range("dataSetPendientes")(1, 7).EntireColumn.NumberFormat = "#,##0.00"
    Else
        MsgBox "No hay Registros"
    End If
    
    closeRS
End Sub

Public Sub btSaldoRango_Click()
    Dim idcuenta As Integer
    Dim i As Integer: i = 0
    
    
    [FECHA1] = ""
    [FECHA2] = ""
    
    frmCalendario2.Show
    If [FECHA1] <> "" And [FECHA2] <> "" Then
        
        OpenDB
        idcuenta = idCuentaPorNumero(rs, cnn, [CUENTA])
        
        If idcuenta = 0 Then
            MsgBox "Error, cuenta no encontrada"
            Exit Sub
        End If
        
        strSQL = "SELECT DB_SALDO_BANCO.FECHA, DB_SALDO_SISTEMA.SALDO_SISTEMA, DB_SALDO_BANCO.SALDO_BANCO FROM DB_SALDO_SISTEMA LEFT JOIN DB_SALDO_BANCO ON DB_SALDO_SISTEMA.FECHA = DB_SALDO_BANCO.FECHA AND DB_SALDO_SISTEMA.ID_CUENTA_FK = DB_SALDO_BANCO.ID_CUENTA_FK WHERE DB_SALDO_BANCO.ID_CUENTA_FK = " & idcuenta & " AND DB_SALDO_BANCO.FECHA >= #" & fechaStrStr([FECHA1]) & "# AND DB_SALDO_BANCO.FECHA <= #" & fechaStrStr([FECHA2]) & "#"
        
        On Error Resume Next
        rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
        On Error GoTo 0
        
        'Log del Error
        logError cnn, Me.Name, "btConsultaSaldo_Click"
        
        Hoja2.Range(Hoja2.Range("dataSetSaldoBanco"), Hoja2.Range("dataSetSaldoBanco").End(xlDown)).ClearContents
        
        If rs.RecordCount > 0 Then
            Hoja2.Activate
            Hoja2.Range("dataSetSaldoBanco").CopyFromRecordset rs
            Hoja2.Range("dataSetSaldoBanco")(1, 2).EntireColumn.NumberFormat = "#,##0.00"
            Hoja2.Range("dataSetSaldoBanco")(1, 3).EntireColumn.NumberFormat = "#,##0.00"
        Else
            MsgBox "No hay Registros"
        End If
        closeRS
    End If
End Sub

Public Sub btGuardar_Click()
    UltimaFunc = "btGuardar_Click"
    UltimaFecha = fechaStrStr(tbFecha.Text)
    Dim idcuenta As Integer
    Dim i As Integer: i = 0
    Dim valido As Boolean: valido = True
    Dim errorFila As String: errorFila = ""
    Dim inicioRango As Range
    
    Set inicioRango = Me.Range("dataSet")(1, 1)
    
    OpenDB
    idcuenta = idCuentaPorNumero(rs, cnn, [CUENTA])
    
    If idcuenta = 0 Then
        MsgBox "Error, cuenta no encontrada"
        Exit Sub
    End If
    
    'Guardar el Saldo del Sistema
    ''determinar si es una actualizacion de saldo
    
    strSQL = "SELECT COUNT(*) FROM DB_SALDO_SISTEMA WHERE ID_CUENTA_FK = " & idcuenta & " AND FECHA = #" & fechaDateStr(UltimaFecha) & "#"
    
    On Error Resume Next
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    'Log del Error
    logError cnn, Me.Name, "btGuardar_Click(1)"
    
    If rs.Fields(0) > 0 Then
        'Es una actualizacion de Saldos
        strSQL = "UPDATE DB_SALDO_SISTEMA SET SALDO_SISTEMA = " & [SALDO_SISTEMA] & " WHERE ID_CUENTA_FK = " & idcuenta & " AND FECHA = #" & fechaDateStr(UltimaFecha) & "#"
    Else
        'Es un nuevo Saldo
        strSQL = "INSERT INTO DB_SALDO_SISTEMA (SALDO_SISTEMA, FECHA, ID_CUENTA_FK) VALUES ( " & [SALDO_SISTEMA] & ", #" & fechaDateStr(UltimaFecha) & "#, " & idcuenta & ")"
    End If
    
    On Error Resume Next
    cnn.Execute strSQL
    On Error GoTo 0
    
    'Log del Error
    logError cnn, Me.Name, "btGuardar_Click(2)"
    
    'Guardar el Saldo del Banco
    ''determinar si es una actualizacion de saldo
    
    strSQL = "SELECT COUNT(*) FROM DB_SALDO_BANCO WHERE ID_CUENTA_FK = " & idcuenta & " AND FECHA = #" & fechaDateStr(UltimaFecha) & "#"
    
    Set rs = Nothing
    On Error Resume Next
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    'Log del Error
    logError cnn, Me.Name, "btGuardar_Click(1)"
    
    If rs.Fields(0) > 0 Then
        'Es una actualizacion de Saldos
        strSQL = "UPDATE DB_SALDO_BANCO SET SALDO_BANCO = " & [SALDO_BANCO] & " WHERE ID_CUENTA_FK = " & idcuenta & " AND FECHA = #" & fechaDateStr(UltimaFecha) & "#"
    Else
        'Es un nuevo Saldo
        strSQL = "INSERT INTO DB_SALDO_BANCO (SALDO_BANCO, FECHA, ID_CUENTA_FK) VALUES ( " & [SALDO_BANCO] & ", #" & fechaDateStr(UltimaFecha) & "#, " & idcuenta & ")"
    End If
    
    On Error Resume Next
    cnn.Execute strSQL
    On Error GoTo 0
    
    'Log del Error
    logError cnn, Me.Name, "btGuardar_Click(2)"
    
    i = 0
    'Loop hasta que las cuatro columnas esten vacias
    
    While inicioRango.Offset(i, 0) <> "" Or inicioRango.Offset(i, 1) <> "" Or inicioRango.Offset(i, 2) <> "" Or inicioRango.Offset(i, 3) <> ""
        If inicioRango.Offset(i, 2) = "" Or inicioRango.Offset(i, 3) = "" Or Not IsNumeric(inicioRango.Offset(i, 0)) Then
            If errorFila = "" Then
                errorFila = inicioRango.Offset(i, 0).row
            Else
                errorFila = errorFila & ", " & inicioRango.Offset(i, 0).row
            End If
            valido = False
        End If
        i = i + 1
    Wend
    
    If Not valido Then
        MsgBox "Error en fila: " & errorFila
    Else
        i = 0
        While inicioRango.Offset(i, 0) <> "" Or inicioRango.Offset(i, 2) <> "" Or inicioRango.Offset(i, 3) <> ""
            If inicioRango.Offset(i, 0) <> "" Then
                'Actualizar
                strSQL = "UPDATE PENDIENTE SET FECHA_INICIO = #" & fechaStrStr(inicioRango.Offset(i, 1)) & "#, IMPORTE = " & inicioRango.Offset(i, 2) & ", TIPO = '" & inicioRango.Offset(i, 3) & "'"
                
                If inicioRango.Offset(i, 4) = "" Then
                    strSQL = strSQL & ", DETALLE = NULL"
                Else
                    strSQL = strSQL & ", DETALLE = '" & inicioRango.Offset(i, 4) & "'"
                End If
                
                If inicioRango.Offset(i, 5) = "" Then
                    strSQL = strSQL & ", ADICIONAL = NULL"
                Else
                    strSQL = strSQL & ", ADICIONAL = '" & inicioRango.Offset(i, 5) & "'"
                End If
                
                strSQL = strSQL & " WHERE ID_CUENTA_FK = " & idcuenta & " AND ID_PENDIENTE = " & inicioRango.Offset(i, 0)
            Else
                'Nuevo
                strSQL = "INSERT INTO PENDIENTE (FECHA_INICIO, IMPORTE, TIPO, DETALLE, ADICIONAL, ID_CUENTA_FK) VALUES (#" & fechaDateStr(UltimaFecha) & "#, " & inicioRango.Offset(i, 2) & ", '" & inicioRango.Offset(i, 3) & "'"
                
                If inicioRango.Offset(i, 4) = "" Then
                    strSQL = strSQL & ", NULL"
                Else
                    strSQL = strSQL & ", '" & inicioRango.Offset(i, 4) & "'"
                End If
                
                If inicioRango.Offset(i, 5) = "" Then
                    strSQL = strSQL & ", NULL"
                Else
                    strSQL = strSQL & ", '" & inicioRango.Offset(i, 5) & "'"
                End If
                
                strSQL = strSQL & ", " & idcuenta & ")"
            End If
    
            On Error Resume Next
            cnn.Execute strSQL
            On Error GoTo 0
            
            'Log del Error
            logError cnn, Me.Name, "btGuardar_Click(3)"
            
            i = i + 1
        Wend
        
        'Finalizar o Anular
        i = 0
        While inicioRango.Offset(i, 0) <> ""
            If inicioRango.Offset(i, 6) = "F" Or inicioRango.Offset(i, 6) = "f" Then
                'Finalizar
                strSQL = "UPDATE PENDIENTE SET FECHA_FIN = #" & fechaDateStr(UltimaFecha) & "# WHERE ID_PENDIENTE = " & inicioRango.Offset(i, 0)
        
                On Error Resume Next
                cnn.Execute strSQL
                On Error GoTo 0
                
                'Log del Error
                logError cnn, Me.Name, "btGuardar_Click(4)"
                
            ElseIf inicioRango.Offset(i, 6) = "A" Or inicioRango.Offset(i, 6) = "a" Then
                'Anular
                strSQL = "UPDATE PENDIENTE SET ANULADO = TRUE WHERE ID_PENDIENTE = " & inicioRango.Offset(i, 0)
        
                On Error Resume Next
                cnn.Execute strSQL
                On Error GoTo 0
                
                'Log del Error
                logError cnn, Me.Name, "btGuardar_Click(5)"
            End If
            
            i = i + 1
        Wend
        
        btActualizar_Click
    End If
End Sub


Public Sub btMasDia_Click()
    UltimaFunc = "btMasDia_Click"
    Dim tmpFecha As String
    
    tmpFecha = fechaStrStr(tbFecha.Text)
    
    If IsDate(tmpFecha) Then
        tmpFecha = DateAdd("d", 1, tmpFecha)
        UltimaFecha = tmpFecha
        tbFecha.Text = Format(Day(tmpFecha), "00") & "/" & Format(Month(tmpFecha), "00") & "/" & Year(tmpFecha)
        btActualizar_Click
    Else
        MsgBox "Error en Fecha"
    End If
End Sub

Public Sub btMenosDia_Click()
    UltimaFunc = "btMasDia_Click"
    Dim tmpFecha As String
    
    tmpFecha = fechaStrStr(tbFecha.Text)
    If IsDate(tmpFecha) Then
        tmpFecha = DateAdd("d", -1, tmpFecha)
        UltimaFecha = tmpFecha
        tbFecha.Text = Format(Day(tmpFecha), "00") & "/" & Format(Month(tmpFecha), "00") & "/" & Year(tmpFecha)
        btActualizar_Click
    Else
        MsgBox "Error en Fecha"
    End If
End Sub

Public Sub btBuscarCuenta_Click()
    UltimaFunc = "btBuscarCuenta_Click"
    Dim tempCuenta As Integer
    If IsNumeric([ID_CUENTA]) Then
        tempCuenta = [ID_CUENTA]
    End If
    busqCuenta.Show
    If [ID_CUENTA] <> tempCuenta Then
        btActualizar_Click
    End If
End Sub

Public Sub btXLSXgenerados_Click()
    busqXLSX.Show
End Sub
