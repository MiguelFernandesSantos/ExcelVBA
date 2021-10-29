'---------------------------------------------
'                                            '
'   Nome: Miguel Fernandes Santos.           '
'   Data: 29/10/2021                         '
'   Projeto: Editanto Planilha.              '
'                                            '
'---------------------------------------------

Private Sub CommandButton1_Click()

' -------- HOLERITE FORMATAÇÃO PARTE 1 --------'

ULTIMA_1 = Sheets("holerites").Range("c7").End(xlDown).Row

If ULTIMA_1 = 1048576 Then

    ULTIMA_1 = 7
    
Else

    ULTIMA_1 = Sheets("holerites").Range("c1048576").End(xlUp).Row + 2
    
End If

If ULTIMA_1 = 7 Then

    ULTIMA_2 = Sheets("holerites").Cells(ULTIMA_1, 3).Offset(13, 4).Row
    
Else

    ULTIMA_2 = Sheets("holerites").Cells(ULTIMA_1, 3).Offset(13, 4).Row
    
End If

f1 = ULTIMA_1 + 1
f2 = ULTIMA_1 + 13

' --- BORDAS --- '

Sheets("holerites").Range("c" & ULTIMA_1 & ":" & "G" & ULTIMA_2).Borders(xlEdgeBottom).LineStyle = 7
Sheets("holerites").Range("c" & ULTIMA_1 & ":" & "G" & ULTIMA_2).Borders(xlEdgeTop).LineStyle = 7
Sheets("holerites").Range("c" & ULTIMA_1 & ":" & "G" & ULTIMA_2).Borders(xlEdgeLeft).LineStyle = 7
Sheets("holerites").Range("c" & ULTIMA_1 & ":" & "G" & ULTIMA_2).Borders(xlEdgeRight).LineStyle = 7

' --- DENTRO --- '
Sheets("holerites").Range("c" & ULTIMA_1 & ":" & "G" & ULTIMA_2).Borders(xlInsideHorizontal).LineStyle = 1
Sheets("holerites").Range("c" & ULTIMA_1 & ":" & "G" & ULTIMA_2).Borders(xlInsideVertical).LineStyle = 1

' --- ESCREVER ---'
Sheets("holerites").Cells(ULTIMA_1, 3).Value = "CODIGO"
Sheets("holerites").Cells(ULTIMA_1, 3).Offset(0, 1).Value = "DESCRIÇÃO"
Sheets("holerites").Cells(ULTIMA_1, 3).Offset(0, 2).Value = "REFERENCIA"
Sheets("holerites").Cells(ULTIMA_1, 3).Offset(0, 3).Value = "VENCIMENTO"
Sheets("holerites").Cells(ULTIMA_1, 3).Offset(0, 4).Value = "DESCONTO"

' --- COLUNA CODIGO --- '
Sheets("holerites").Cells(ULTIMA_1, 3).Offset(1, 0).Value = "1"
Sheets("holerites").Cells(ULTIMA_1, 3).Offset(2, 0).Value = "2"
Sheets("holerites").Cells(ULTIMA_1, 3).Offset(3, 0).Value = "3"
Sheets("holerites").Cells(ULTIMA_1, 3).Offset(4, 0).Value = "4"
Sheets("holerites").Cells(ULTIMA_1, 3).Offset(5, 0).Value = "5"
Sheets("holerites").Cells(ULTIMA_1, 3).Offset(6, 0).Value = "6"
Sheets("holerites").Cells(ULTIMA_1, 3).Offset(7, 0).Value = "7"
Sheets("holerites").Cells(ULTIMA_1, 3).Offset(8, 0).Value = "8"
Sheets("holerites").Cells(ULTIMA_1, 3).Offset(9, 0).Value = "9"
Sheets("holerites").Cells(ULTIMA_1, 3).Offset(10, 0).Value = "10"
Sheets("holerites").Cells(ULTIMA_1, 3).Offset(11, 0).Value = "11"
Sheets("holerites").Cells(ULTIMA_1, 3).Offset(12, 0).Value = "12"
Sheets("holerites").Cells(ULTIMA_1, 3).Offset(13, 0).Value = "13"

' -------- HOLERITE FORMATAÇÃO PARTE 2 --------'
ULTIMA_1 = Sheets("holerites").Range("c1048576").End(xlUp).Row + 1
ULTIMA_2 = Sheets("holerites").Range("c1048576").End(xlUp).Row + 5

' --- BORDAS --- '
Sheets("holerites").Range("c" & ULTIMA_1 & ":" & "G" & ULTIMA_2).Borders(xlEdgeBottom).LineStyle = 7
Sheets("holerites").Range("c" & ULTIMA_1 & ":" & "G" & ULTIMA_2).Borders(xlEdgeTop).LineStyle = 7
Sheets("holerites").Range("c" & ULTIMA_1 & ":" & "G" & ULTIMA_2).Borders(xlEdgeLeft).LineStyle = 7
Sheets("holerites").Range("c" & ULTIMA_1 & ":" & "G" & ULTIMA_2).Borders(xlEdgeRight).LineStyle = 7

' --- DENTRO --- '
Sheets("holerites").Range("c" & ULTIMA_1 & ":" & "G" & ULTIMA_2).Borders(xlInsideHorizontal).LineStyle = 1
Sheets("holerites").Range("c" & ULTIMA_1 & ":" & "G" & ULTIMA_2).Borders(xlInsideVertical).LineStyle = 1

' --------- SALARIO LIQUIDO --------- '
Sheets("holerites").Cells(ULTIMA_1, 5).Value = "TOTAL VENCIMENTOS ->"

Sheets("holerites").Cells(ULTIMA_1, 3).Offset(1, 0).Value = "DATA ->"
Sheets("holerites").Cells(ULTIMA_1, 3).Offset(1, 1).Value = "    /    /    "

Sheets("holerites").Cells(ULTIMA_1, 3).Offset(1, 2).Value = "TOTAL DESCONTO ->"

Sheets("holerites").Cells(ULTIMA_1, 3).Offset(2, 2).Value = "SALARIO LIQUIDO ->"

' --------- IMPOSTOS --------- '
Sheets("holerites").Cells(ULTIMA_1, 3).Offset(3, 0).Value = "BASE INSS"
Sheets("holerites").Cells(ULTIMA_1, 3).Offset(3, 1).Value = "BASE FGTS"
Sheets("holerites").Cells(ULTIMA_1, 3).Offset(3, 2).Value = "FGTS MÊS"
Sheets("holerites").Cells(ULTIMA_1, 3).Offset(3, 3).Value = "BASE IRRF"
Sheets("holerites").Cells(ULTIMA_1, 3).Offset(3, 4).Value = "FAIXA IRRF"

'FIM DO CODIGO :)

End Sub
