Private Sub Workbook_Open()
    UltimaFecha = Now()
    If Format(DateAdd("d", -1, UltimaFecha), "DDDD") = "domingo" Then
        UltimaFecha = DateAdd("d", -2, UltimaFecha)
        Hoja1.tbFecha.Text = Format(Day(UltimaFecha), "00") & "/" & Format(Month(UltimaFecha), "00") & "/" & Year(UltimaFecha)
    Else
        UltimaFecha = DateAdd("d", -1, UltimaFecha)
        Hoja1.tbFecha.Text = Format(Day(UltimaFecha), "00") & "/" & Format(Month(UltimaFecha), "00") & "/" & Year(UltimaFecha)
    End If
    Hoja1.btActualizar_Click
End Sub
