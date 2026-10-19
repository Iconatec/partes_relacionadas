Sub nd()
Dim ws As Worksheet
Dim ws2 As Worksheet
Set ws = ActiveWorkbook.Sheets("Ativo")
Set ws2 = ActiveWorkbook.Sheets("Passivo")
Dim ultimaLinha As Long
Dim i As Long

'GUIA ATIVO
ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

With ws
.Rows(ultimaLinha).Delete
.Rows(ultimaLinha - 1).Delete
.Rows(1).Delete
.Columns("B:E").Delete
.Columns("A:C").Insert
.Range("A1").Value = "Ordem"
.Range("B1").Value = "Filial"
.Range("C1").Value = "Agente"
.Range("E1").Value = "Saldo anterior"
.Range("F1").Value = "Débito"
.Range("G1").Value = "Crédito"
.Range("H1").Value = "Saldo"
.Range("I1").Value = "Concatenar"
.Range("J1").Value = "Passivo"
.Range("K1").Value = "Dif"
.Range("L1").Value = "Obs"
End With

'Preenchimento das colunas filial e agente
ultimaLinha = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
a = 1
For i = 2 To ultimaLinha
    If ws.Cells(i, "D").Value Like "Filial:*" Then
        ws.Cells(i, "B").Formula = "=textbefore(textafter(D" & i & ", ""Filial: ""),"" -"")"
        ws.Cells(i, "B").Value = ws.Cells(i, "B").Value
        ws.Rows(i).Interior.Color = RGB(200, 200, 200)
    ElseIf ws.Cells(i, "D").Value Like "Agente:*" Then
        ws.Cells(i, "C").Formula = "=textbefore(textafter(D" & i & ", ""Agente: ""),"" -"")"
        ws.Cells(i, "C").Value = ws.Cells(i, "C").Value
        ws.Cells(i, "I").NumberFormat = "General"
        ws.Cells(i, "I").Formula = "=CONCAT(B" & i & ",""-"",C" & i & ")"
        
        ws.Cells(i, "K").Formula = "=H" & i & " + J" & i
    End If
    
'Coluna ordem
ws.Cells(i, "A").Value = a
a = a + 1

'Preencher a filial
If ws.Cells(i, "B").Value = "" Then
    ws.Cells(i, "B").Value = ws.Cells(i - 1, "B").Value
End If
ws.Cells(i, "I").NumberFormat = "@"
ws.Cells(i, "I").Value = ws.Cells(i, "I").Value
Next i

'GUIA PASSIVO
ultimaLinha = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row

With ws2
.Rows(ultimaLinha).Delete
.Rows(ultimaLinha - 1).Delete
.Rows(1).Delete
.Columns("B:E").Delete
.Columns("A:C").Insert
.Range("A1").Value = "Ordem"
.Range("B1").Value = "Filial"
.Range("C1").Value = "Agente"
.Range("E1").Value = "Saldo anterior"
.Range("F1").Value = "Débito"
.Range("G1").Value = "Crédito"
.Range("H1").Value = "Saldo"
.Range("I1").Value = "Concatenar"
.Range("J1").Value = "Ativo"
.Range("K1").Value = "Dif"
.Range("L1").Value = "Obs"
End With

ultimaLinha = ws2.Cells(ws2.Rows.Count, "D").End(xlUp).Row
a = 1
For i = 2 To ultimaLinha
    If ws2.Cells(i, "D").Value Like "Filial:*" Then
        ws2.Cells(i, "B").Formula = "=textbefore(textafter(D" & i & ", ""Filial: ""),"" -"")"
        ws2.Cells(i, "B").Value = ws2.Cells(i, "B").Value
        ws2.Rows(i).Interior.Color = RGB(200, 200, 200)
    ElseIf ws2.Cells(i, "D").Value Like "Agente:*" Then
        ws2.Cells(i, "C").Formula = "=textbefore(textafter(D" & i & ", ""Agente: ""),"" -"")"
        ws2.Cells(i, "C").Value = ws2.Cells(i, "C").Value
        ws2.Cells(i, "I").NumberFormat = "General"
        ws2.Cells(i, "I").Formula = "=CONCAT(C" & i & ",""-"",B" & i & ")"
        ws2.Cells(i, "J").Formula = "=XLOOKUP(I" & i & ",Ativo!I:I,Ativo!H:H)"
        ws2.Cells(i, "K").Formula = "=H" & i & " + J" & i
    End If
    
    'Coluna ordem
    ws2.Cells(i, "A").Value = a
    a = a + 1

    'Preencher a filial
    If ws2.Cells(i, "B").Value = "" Then
        ws2.Cells(i, "B").Value = ws2.Cells(i - 1, "B").Value
    End If
    ws2.Cells(i, "I").NumberFormat = "@"
    ws2.Cells(i, "I").Value = ws2.Cells(i, "I").Value
Next i

'Arrumar o procx do ativo
ultimaLinha = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
For i = 2 To ultimaLinha
If ws.Cells(i, "D").Value Like "Agente:*" Then
ws.Cells(i, "J").Formula = "=XLOOKUP(I" & i & ",Passivo!I:I,Passivo!H:H)"
End If
Next i

'Conferência das diferenças no ativo
ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
For i = 2 To ultimaLinha
    If IsError(ws.Cells(i, "K").Value) Then
        ws.Rows(i).Interior.Color = RGB(255, 192, 0)
    ElseIf ws.Cells(i, "K").Value <> 0 Then
        ws.Rows(i).Interior.Color = RGB(255, 192, 0)
    End If
Next i

'Conferência das diferenças no passivo
ultimaLinha = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
For i = 2 To ultimaLinha
    If IsError(ws2.Cells(i, "K").Value) Then
        ws2.Rows(i).Interior.Color = RGB(255, 192, 0)
    ElseIf ws2.Cells(i, "K").Value <> 0 Then
        ws2.Rows(i).Interior.Color = RGB(255, 192, 0)
    End If
Next i

'Autofit nas colunas
ws.Columns("A:K").AutoFit
ws2.Columns("A:K").AutoFit
'Congelar linha 1
ws.Activate
ws.Rows("2:2").Select
ActiveWindow.FreezePanes = True
ws2.Activate
ws2.Rows("2:2").Select
ActiveWindow.FreezePanes = True
'Filtro na linha 1
ws.Rows(1).AutoFilter
ws2.Rows(1).AutoFilter

'Contagem de divergências ATIVO
contadorWS = 0
ultimaLinhaWS = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
For i = 2 To ultimaLinhaWS
    If IsError(ws.Cells(i, "K").Value) Then
        contadorWS = contadorWS + 1
    ElseIf ws.Cells(i, "K").Value <> 0 Then
        contadorWS = contadorWS + 1
    End If
Next i

'Contagem de divergências PASSIVO
contadorWS2 = 0
ultimaLinhaWS2 = ws2.Cells(ws2.Rows.Count, "K").End(xlUp).Row
For i = 2 To ultimaLinhaWS2
    If IsError(ws2.Cells(i, "K").Value) Then
        contadorWS2 = contadorWS2 + 1
    ElseIf ws2.Cells(i, "K").Value <> 0 Then
        contadorWS2 = contadorWS2 + 1
    End If
Next i

MsgBox "Total de divergências:" & vbCrLf & _
       "Ativo: " & contadorWS & vbCrLf & _
       "Passivo: " & contadorWS2, vbInformation, "Contagem de erros"

End Sub
