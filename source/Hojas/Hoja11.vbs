Private Sub btComparar_Click()

    Application.Calculation = xlManual
    Application.ScreenUpdating = False
        
    Range(Range("LISTA_A2"), Cells(1000, Range("LISTA_B2").Column)).ClearContents
    
    ActiveWorkbook.Worksheets("Comparar Lista").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Comparar Lista").Sort.SortFields.Add Key:=Range( _
        "LISTA_A2"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
        
    Range("LISTA_A").CurrentRegion.Copy
    Range("LISTA_A2").PasteSpecial xlPasteValues
    With ActiveWorkbook.Worksheets("Comparar Lista").Sort
        .SetRange Range("LISTA_A2").CurrentRegion
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ActiveWorkbook.Worksheets("Comparar Lista").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Comparar Lista").Sort.SortFields.Add Key:=Range( _
        "LISTA_B2"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
        
    Range("LISTA_B").CurrentRegion.Copy
    Range("LISTA_B2").PasteSpecial xlPasteValues
    With ActiveWorkbook.Worksheets("Comparar Lista").Sort
        .SetRange Range("LISTA_B2").CurrentRegion
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Dim i As Integer
    i = 0
    While Range("LISTA_A2").Offset(i, 0).Value <> ""
        If Range("LISTA_A2").Offset(i, 0).Value < Range("LISTA_B2").Offset(i, 0).Value Then
            Range("LISTA_A2").Offset(i, 0).Insert Shift:=xlDown
        ElseIf Range("LISTA_A2").Offset(i, 0).Value > Range("LISTA_B2").Offset(i, 0).Value Then
            Range("LISTA_B2").Offset(i, 0).Insert Shift:=xlDown
        End If
        i = i + 1
    Wend
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Private Sub btLimpiar_Click()
    Range(Range("LISTA_A"), Cells(1000, Range("LISTA_B").Column)).ClearContents
    Range(Range("LISTA_A2"), Cells(1000, Range("LISTA_B2").Column)).ClearContents
End Sub

Private Sub btRegresar_Click()
    Hoja1.Activate
End Sub
