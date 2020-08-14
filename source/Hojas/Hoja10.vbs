
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
            If RealEqual(CurrTotal + InArr(i, 1), TargetVal, Epsilon) Then
                Rslt(UBound(Rslt)) = ExtendRslt(CurrRslt, i, Separator)
                If MaxSoln = 0 Then
                    If UBound(Rslt) Mod 100 = 0 Then Debug.Print UBound(Rslt) & "=" & Rslt(UBound(Rslt))
                Else
                    If UBound(Rslt) >= MaxSoln Then Exit Sub
                End If
                ReDim Preserve Rslt(UBound(Rslt) + 1)
            ElseIf CurrTotal + InArr(i, 1) > TargetVal + Epsilon Then
            ElseIf CurrIdx < UBound(InArr) Then
                recursiveMatch MaxSoln, TargetVal, InArr(), i + 1, _
                    CurrTotal + InArr(i, 1), Epsilon, Rslt(), _
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
            If RealEqual(CurrTotal + InArr(i, 1), TargetVal, Epsilon) Then
                Rslt(UBound(Rslt)) = ExtendRslt(CurrRslt, i, Separator)
                If MaxSoln = 0 Then
                    If UBound(Rslt) Mod 100 = 0 Then Debug.Print UBound(Rslt) & "=" & Rslt(UBound(Rslt))
                Else
                    If UBound(Rslt) >= MaxSoln Then Exit Sub
                End If
                ReDim Preserve Rslt(UBound(Rslt) + 1)
            ElseIf CurrTotal + InArr(i, 1) < TargetVal + Epsilon Then
            ElseIf CurrIdx < UBound(InArr) Then
                recursiveMatch MaxSoln, TargetVal, InArr(), i + 1, _
                    CurrTotal + InArr(i, 1), Epsilon, Rslt(), _
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

Private Sub btBuscarSuma_Click()
    If [OBJETIVO] <> "" And [VALORES] <> "" Then
        Me.Range(Me.Range("ALTERNATIVA"), Me.Range("ALTERNATIVA").Offset(1000, 0)).ClearContents
        If IsNumeric([OBJETIVO]) And IsNumeric([VALORES]) Then
            Dim Rslt(), InArr() As Variant, StartTime As Date, MaxSoln As Integer
            InArr = Me.Range("VALORES").CurrentRegion
            Me.Range("POSIBILIDADES").CurrentRegion.ClearContents
            StartTime = Now()
            MaxSoln = 0
            ReDim Rslt(0)
            recursiveMatch MaxSoln, CDbl(Me.Range("OBJETIVO")), InArr, LBound(InArr), 0, 0.00000001, _
                Rslt, "", ";"
            Dim idxMatch As Variant
            
            Dim rowAnswer As Variant
            cmbAlternativas.Clear
            
            For i = 0 To UBound(Rslt) - 1
                cmbAlternativas.AddItem Rslt(i)
                rowAnswer = Split(Rslt(i), ";")
                For j = 0 To UBound(rowAnswer)
                    Me.Range("POSIBILIDADES").Offset(j, i) = InArr(rowAnswer(j), 1)
                Next j
            Next i
            MsgBox "Numero de Resultados: " & UBound(Rslt)
        End If
    End If
End Sub

Private Sub btRegresar_Click()
    Hoja1.Activate
End Sub

Private Sub cmbAlternativas_Change()
    If cmbAlternativas.ListIndex <> -1 Then
        Me.Range(Me.Range("ALTERNATIVA"), Me.Range("ALTERNATIVA").Offset(1000, 0)).ClearContents
        
        Dim rowAnswer As Variant
        rowAnswer = Split(cmbAlternativas.List(cmbAlternativas.ListIndex), ";")
        
        For j = 0 To UBound(rowAnswer)
            Me.Range("ALTERNATIVA").Offset(rowAnswer(j) - 1, 0) = cmbAlternativas.ListIndex + 1
        Next j
    End If
End Sub
