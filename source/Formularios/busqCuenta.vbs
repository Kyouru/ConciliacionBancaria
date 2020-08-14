
Private Sub btAceptar_Click()
    If ListBox1.ListIndex <> -1 Then
        [ID_CUENTA] = ListBox1.List(ListBox1.ListIndex, 0)
        [BANCO] = ListBox1.List(ListBox1.ListIndex, 1)
        [CUENTA] = ListBox1.List(ListBox1.ListIndex, 2)
        [MONEDA] = ListBox1.List(ListBox1.ListIndex, 3)
        Unload Me
    End If
End Sub

Private Sub btCancelar_Click()
    Unload Me
End Sub

Private Sub cmbBanco_Change()
    If cmbBanco.ListIndex <> -1 Then
        ActualizarHoja
        ActualizarLista
    End If
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btAceptar_Click
End Sub

Private Sub obDolares_Change()
    obSoles_Change
End Sub

Private Sub obSoles_Change()
    ActualizarHoja
    ActualizarLista
End Sub

Private Sub UserForm_Initialize()
    Dim cont As Integer
    strSQL = "SELECT * FROM BANCO"
    OpenDB
    'On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    
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
    
    ActualizarHoja
    ActualizarLista
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - UserForm_Initialize", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub


'Se Solicita todas las Condiciones del Prestamo Seleccionado previamente y se Copian a una hoja Temporal para luego poder agregarlos a la ListBox
Public Sub ActualizarHoja()
    strSQL = "SELECT ID_CUENTA, NOMBRE_BANCO, NUMERO_CUENTA_SISGO, NOMBRE_MONEDA" & _
            " FROM ((CUENTA AS C LEFT JOIN BANCO AS B ON C.ID_BANCO_FK = B.ID_BANCO)" & _
            " LEFT JOIN MONEDA AS M ON M.ID_MONEDA = C.ID_MONEDA_FK)" & _
            " WHERE 1=1 "
    If cmbBanco.ListIndex <> -1 Then
        strSQL = strSQL & " AND C.ID_BANCO_FK = " & cmbBanco.List(cmbBanco.ListIndex, 1)
    End If
    If obSoles Then
        strSQL = strSQL & " AND C.ID_MONEDA_FK = " & ID_SOLES
    Else
        If obDolares Then
            strSQL = strSQL & " AND C.ID_MONEDA_FK = " & ID_DOLARES
        End If
    End If
    
    strSQL = strSQL & " ORDER BY NOMBRE_BANCO, NOMBRE_MONEDA DESC, NUMERO_CUENTA_SISGO"
    
    With ThisWorkbook.Sheets(CStr([TEMP_CUENTA]))
        'Limpiar Hoja Temporal
        .Range(.Range([DATASET_TEMP_CUENTA]), .Range([DATASET_TEMP_CUENTA]).End(xlDown)).ClearContents
        
        OpenDB
        On Error GoTo Handle:
        rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount > 0 Then
            .Range([DATASET_TEMP_CUENTA]).CopyFromRecordset rs
        End If
        closeRS
    End With
    
Handle:
    If cnn.Errors.count > 0 Then
        'Log del Error
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - ActualizarHoja", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

'Agrega la Hoja Temporal a la ListBox
Public Sub ActualizarLista()
    With ThisWorkbook.Sheets(CStr([TEMP_CUENTA]))
        ListBox1.ColumnWidths = "0;150;100;0;"
        ListBox1.ColumnCount = 4
        ListBox1.ColumnHeads = True
        'En caso halla mas de una fila
        If .Range("A3") <> "" Then
            ListBox1.RowSource = .Name & "!" & Left(.Range([DATASET_TEMP_CUENTA]).Address, Len(.Range([DATASET_TEMP_CUENTA]).Address) - 1) & .Range("A3").End(xlDown).row
        Else
            'En caso halla solamente una fila
            If .Range("A2") <> "" Then
                ListBox1.RowSource = .Name & "!" & .Range([DATASET_TEMP_CUENTA]).Address
            'En caso no hallan datos
            Else
                ListBox1.RowSource = ""
                ListBox1.ColumnHeads = False
            End If
        End If
    End With
End Sub

