
Private Sub btCalendario_Click()
    frmCalendario.Show
    If Not IsNull(Hoja1.UltimaFecha) Then
        tbFecha.Text = Format(Hoja1.UltimaFecha, "DD/MM/YYYY")
        ActualizarHoja
        ActualizarLista
    End If
End Sub

Private Sub cmbBanco_Change()
    If cmbBanco.ListIndex <> -1 Then
    
        strSQL = "SELECT * FROM CUENTA WHERE ID_BANCO_FK = " & cmbBanco.List(cmbBanco.ListIndex, 1)
        OpenDB
        
        On Error Resume Next
        rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
        On Error GoTo 0
        
        'Log del Error
        logError cnn, Me.Name, "cmbBanco_Change"
        
        If rs.RecordCount > 0 Then
            cmbCuenta.Clear
            cont = 0
            Do While Not rs.EOF
                cmbCuenta.AddItem rs.Fields("NUMERO_CUENTA_SISGO")
                cmbCuenta.List(cont, 1) = rs.Fields("ID_CUENTA")
                cmbCuenta.List(cont, 2) = rs.Fields("ID_MONEDA_FK")
                cont = cont + 1
                rs.MoveNext
            Loop
        End If
        
        ActualizarHoja
        ActualizarLista
    End If
End Sub

Private Sub cmbCuenta_Change()
    If cmbCuenta.ListIndex <> -1 Then
        If cmbCuenta.List(cmbCuenta.ListIndex, 2) = 1 Then
            If obSoles Then
                ActualizarHoja
                ActualizarLista
            End If
            obSoles.Enabled = True
            obSoles = True
            obDolares.Enabled = False
        Else
            If obDolares Then
                ActualizarHoja
                ActualizarLista
            End If
            obDolares.Enabled = True
            obDolares = True
            obSoles.Enabled = False
        End If
    Else
        obSoles.Enabled = True
        obDolares.Enabled = True
    End If
End Sub

Private Sub listPendientes_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If listPendientes.ListIndex <> -1 Then
        Hoja1.idPendiente = listPendientes.List(listPendientes.ListIndex, 0)
        frmEditarPendiente.Show
    End If
End Sub

Private Sub obDolares_Change()
    If obDolares Then
        ActualizarHoja
        ActualizarLista
    End If
End Sub

Private Sub obSoles_Change()
    If obSoles Then
        ActualizarHoja
        ActualizarLista
    End If
End Sub

Private Sub tbFecha_Change()
    If tbFecha.Text <> "" Then
        If IsDate(tbFecha.Text) Then
            ActualizarHoja
            ActualizarLista
        End If
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim cont As Integer
    strSQL = "SELECT * FROM BANCO"
    OpenDB
    
    On Error Resume Next
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    'Log del Error
    logError cnn, Me.Name, "UserForm_Initialize"
    
    If rs.RecordCount > 0 Then
        cmbBanco.Clear
        cont = 0
        Do While Not rs.EOF
            cmbBanco.AddItem rs.Fields("NOMBRE_BANCO")
            cmbBanco.List(cont, 1) = rs.Fields("ID_BANCO")
            cont = cont + 1
            rs.MoveNext
        Loop
    End If
    
    strSQL = "SELECT * FROM CUENTA"
    OpenDB
    
    On Error Resume Next
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    'Log del Error
    logError cnn, Me.Name, "UserForm_Initialize"
    
    If rs.RecordCount > 0 Then
        cmbCuenta.Clear
        cont = 0
        Do While Not rs.EOF
            cmbCuenta.AddItem rs.Fields("NUMERO_CUENTA_SISGO")
            cmbCuenta.List(cont, 1) = rs.Fields("ID_CUENTA")
            cmbCuenta.List(cont, 2) = rs.Fields("ID_MONEDA_FK")
            cont = cont + 1
            rs.MoveNext
        Loop
    End If
    
    ActualizarHoja
    ActualizarLista
    
End Sub

'Se Solicita todos los Pendientes que cumplan los filtros y se Pega en una hoja Temporal para luego poder agregarlos a la ListBox
''Es mas agil exportar todo el recordset y pegarlo en una hoja que recorrer el recordset he ir agregando cada elemento
Public Sub ActualizarHoja()

    strSQL = "SELECT ID_PENDIENTE, NOMBRE_BANCO, NUMERO_CUENTA_SISGO, NOMBRE_MONEDA, FECHA_INICIO, FECHA_FIN, " & _
    "IMPORTE, TIPO, DETALLE, ADICIONAL, ANULADO " & _
    "FROM ((PENDIENTE LEFT JOIN CUENTA ON PENDIENTE.ID_CUENTA_FK = CUENTA.ID_CUENTA)" & _
    " LEFT JOIN MONEDA ON MONEDA.ID_MONEDA = CUENTA.ID_MONEDA_FK)" & _
    " LEFT JOIN BANCO ON BANCO.ID_BANCO = CUENTA.ID_BANCO_FK" & _
    " WHERE 1 = 1"
    
    If cmbBanco.ListIndex <> -1 Then
        strSQL = strSQL & " AND ID_BANCO = " & cmbBanco.List(cmbBanco.ListIndex, 1)
    End If
    
    If obSoles Then
        strSQL = strSQL & " AND ID_MONEDA_FK = 1"
    Else
        If obDolares Then
            strSQL = strSQL & " AND ID_MONEDA_FK = 2"
        End If
    End If
    
    If cmbCuenta.ListIndex <> -1 Then
        strSQL = strSQL & " AND ID_CUENTA = " & cmbCuenta.List(cmbCuenta.ListIndex, 1)
    End If
    
    If tbFecha.Text <> "" Then
        If IsDate(tbFecha.Text) Then
            strSQL = strSQL & " AND FECHA_INICIO = #" & fechaStrStr(tbFecha.Text) & "#"
        End If
    End If
    
    strSQL = strSQL & " ORDER BY FECHA_INICIO DESC, NOMBRE_BANCO, NUMERO_CUENTA_SISGO, FECHA_FIN DESC"
    
    'Limpia Hoja Temporal
    ThisWorkbook.Sheets("HOJA_TEMP_PENDIENTE").Range(ThisWorkbook.Sheets("HOJA_TEMP_PENDIENTE").Range("dataSetPendiente"), ThisWorkbook.Sheets("HOJA_TEMP_PENDIENTE").Range("dataSetPendiente").End(xlDown)).ClearContents
    
    OpenDB
    
    On Error Resume Next
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    'Log del Error
    logError cnn, Me.Name, "ActualizarHoja"
    
    If rs.RecordCount > 0 Then
        ThisWorkbook.Sheets("HOJA_TEMP_PENDIENTE").Range("dataSetPendiente").CopyFromRecordset rs
    End If
    
End Sub


'Agrega la Hoja Temporal a la ListBox
Public Sub ActualizarLista()
    With ThisWorkbook.Sheets("HOJA_TEMP_PENDIENTE")
        .Range("E:F").NumberFormat = "DD/MM/YYYY"
        .Range("G:G").NumberFormat = "#,###,#0.00"
        
        listPendientes.ColumnWidths = "0;80;120;40;60;60;80;80;80;80;80"
        listPendientes.ColumnCount = 11
        listPendientes.ColumnHeads = True
        'En caso halla mas de una fila
        If .Range("A3") <> "" Then
            listPendientes.RowSource = .Name & "!" & Left(.Range("dataSetPendiente").Address, Len(.Range("dataSetPendiente").Address) - 1) & .Range("A2").End(xlDown).row
        Else
            'En caso halla solamente una fila
            If .Range("A2") <> "" Then
                listPendientes.RowSource = .Name & "!" & .Range("dataSetPendiente").Address
            'En caso no hallan datos
            Else
                listPendientes.RowSource = ""
                listPendientes.ColumnHeads = False
            End If
        End If
        
    End With
    
End Sub


