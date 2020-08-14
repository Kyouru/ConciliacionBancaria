
Private Sub btBuscar_Click()
    Dim Rslt(), InArr(), StartTime As Date, MaxSoln As Integer
    Dim rango As Range
    Dim splitRes As Variant
    Dim tmpRes As String
    
    Set rango = Hoja1.Range(Hoja1.Range("dataSet")(1, 3), Hoja1.Range("dataSet")(1, 3).End(xlDown))
    
    StartTime = Now()
    MaxSoln = 0
    For i = 1 To rango.Rows.count
        ReDim Preserve InArr(i - 1)
        InArr(i - 1) = Abs(rango(i, 1))
    Next i
    ReDim Rslt(0)
    recursiveMatch MaxSoln, CDbl(tbSuma.Text), InArr, LBound(InArr), 0, 0.00000001, _
        Rslt, "", ";"
    Dim idxMatch As Variant
    
    For i = 0 To UBound(Rslt) - 1
        tmpRes = ""
        splitRes = Split(Rslt(i), ";")
        For j = 0 To UBound(splitRes)
            tmpRes = tmpRes & splitRes(j) + Hoja1.Range("dataSet").row & ";"
        Next j
        cmbResultado.AddItem tmpRes
    Next i
    'MsgBox UBound(Rslt) & " Resultados"
    If cmbResultado.ListCount > 0 Then
        cmbResultado.ListIndex = 0
    Else
        MsgBox "No se encontr・Combinaci Positiva"
    End If
End Sub

Function RealEqual(A, B, Epsilon As Double)
    RealEqual = Abs(A - B) <= Epsilon
End Function

Function ExtendRslt(CurrRslt, NewVal, Separator)
    If CurrRslt = "" Then ExtendRslt = NewVal _
    Else ExtendRslt = CurrRslt & Separator & NewVal
End Function

Sub recursiveMatch(ByVal MaxSoln As Integer, ByVal TargetVal As Double, InArr(), _
        ByVal CurrIdx As Integer, _
        ByVal CurrTotal, ByVal Epsilon As Double, _
        ByRef Rslt(), ByVal CurrRslt As String, ByVal Separator As String)
    Dim i As Integer
    If TargetVal >= 0 Then
        For i = CurrIdx To UBound(InArr)
            If RealEqual(CurrTotal + InArr(i), TargetVal, Epsilon) Then
                Rslt(UBound(Rslt)) = ExtendRslt(CurrRslt, i, Separator)
                If MaxSoln = 0 Then
                    If UBound(Rslt) Mod 100 = 0 Then Debug.Print UBound(Rslt) & "=" & Rslt(UBound(Rslt))
                Else
                    If UBound(Rslt) >= MaxSoln Then Exit Sub
                End If
                ReDim Preserve Rslt(UBound(Rslt) + 1)
            ElseIf CurrTotal + InArr(i) > TargetVal + Epsilon Then
            ElseIf CurrIdx < UBound(InArr) Then
                recursiveMatch MaxSoln, TargetVal, InArr(), i + 1, _
                    CurrTotal + InArr(i), Epsilon, Rslt(), _
                    ExtendRslt(CurrRslt, i, Separator), _
                    Separator
                If MaxSoln <> 0 Then If UBound(Rslt) >= MaxSoln Then Exit Sub
                
                Else
                'we've run out of possible elements and we _
                 still don't have a match
            End If
        Next i
    Else
        For i = CurrIdx To UBound(InArr)
            If RealEqual(CurrTotal + InArr(i), TargetVal, Epsilon) Then
                Rslt(UBound(Rslt)) = ExtendRslt(CurrRslt, i, Separator)
                If MaxSoln = 0 Then
                    If UBound(Rslt) Mod 100 = 0 Then Debug.Print UBound(Rslt) & "=" & Rslt(UBound(Rslt))
                Else
                    If UBound(Rslt) >= MaxSoln Then Exit Sub
                End If
                ReDim Preserve Rslt(UBound(Rslt) + 1)
            ElseIf CurrTotal + InArr(i) < TargetVal + Epsilon Then
            ElseIf CurrIdx < UBound(InArr) Then
                recursiveMatch MaxSoln, TargetVal, InArr(), i + 1, _
                    CurrTotal + InArr(i), Epsilon, Rslt(), _
                    ExtendRslt(CurrRslt, i, Separator), _
                    Separator
                If MaxSoln <> 0 Then If UBound(Rslt) >= MaxSoln Then Exit Sub
                
                Else
                'we've run out of possible elements and we _
                 still don't have a match
            End If
        Next i

    End If
End Sub

Private Sub cmbResultado_Change()
    If cmbResultado.ListIndex <> -1 Then
        Dim rango As Range
        Dim splitRes As Variant
        Dim rowSelect As String
        rowSelect = ""
        
        Set rango = Hoja1.Range(Hoja1.Range("dataSet")(1, 3), Hoja1.Range("dataSet")(1, 3).End(xlDown))
        splitRes = Split(cmbResultado.List(cmbResultado.ListIndex), ";")
        For i = 0 To UBound(splitRes) - 1
            rowSelect = rowSelect & "D" & splitRes(i) & ","
        Next i
        rowSelect = Left(rowSelect, Len(rowSelect) - 1)
        Hoja1.Range(rowSelect).Select
    End If
End Sub
