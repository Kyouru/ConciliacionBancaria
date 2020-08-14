Private Sub btCalendarioInicio_Click()
    
End Sub

Private Sub btCancelar_Click()
    Unload Me
End Sub

Private Sub btGuardar_Click()
    strSQL = "UPDATE PENDIENTE SET " & _
                " FECHA_INICIO = #" & fechaStrStr(tbFechaInicio.Text) & "#" & _
                ", FECHA_FIN = #" & fechaStrStr(tbFechaFin.Text) & "#" & _
                ", IMPORTE = " & tbMonto.Text & _
                ", TIPO = '" & tbTipo.Text & "'"
    
    If tbDetalle.Text <> "" Then
        strSQL = strSQL & ", DETALLE = '" & tbDetalle.Text & "'"
    Else
        strSQL = strSQL & ", DETALLE = NULL"
    End If
    
    If tbAdicional.Text <> "" Then
        strSQL = strSQL & ", ADICIONAL = '" & tbAdicional.Text & "'"
    Else
        strSQL = strSQL & ", ADICIONAL = NULL"
    End If
    
    strSQL = strSQL & " WHERE ID_PENDIENTE = " & lbIdPendiente.Caption
    closeRS
    OpenDB
    
    On Error Resume Next
    cnn.Execute strSQL
    On Error GoTo 0
    
    'Log del Error
    logError cnn, Me.Name, "UserForm_Initialize"
    
    Unload Me
    
    frmMantenimientoPendientes.ActualizarHoja
    frmMantenimientoPendientes.ActualizarLista
    
End Sub

Private Sub UserForm_Initialize()

    strSQL = "SELECT * FROM PENDIENTE WHERE ID_PENDIENTE = " & Hoja1.idPendiente
    OpenDB
    
    On Error Resume Next
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    On Error GoTo 0
    
    'Log del Error
    logError cnn, Me.Name, "UserForm_Initialize"
    
    If rs.RecordCount > 0 Then
        lbIdPendiente.Caption = rs.Fields("ID_PENDIENTE")
        tbFechaInicio.Text = Format(rs.Fields("FECHA_INICIO"), "DD/MM/YYYY")
        tbFechaFin.Text = Format(rs.Fields("FECHA_FIN"), "DD/MM/YYYY")
        tbMonto.Text = rs.Fields("IMPORTE")
        
        If rs.Fields("ANULADO") Then
            cbAnulado = True
        Else
            cbAnulado = False
        End If
        
        If Not IsNull(rs.Fields("TIPO")) Then
            tbTipo.Text = rs.Fields("TIPO")
        End If
        
        If Not IsNull(rs.Fields("DETALLE")) Then
            tbDetalle.Text = rs.Fields("DETALLE")
        End If
        
        If Not IsNull(rs.Fields("ADICIONAL")) Then
            tbAdicional.Text = rs.Fields("ADICIONAL")
        End If
    End If
End Sub
