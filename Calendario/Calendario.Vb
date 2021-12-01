Sub NovoCalendario()
    
    ' variaveis para tarefas
    Dim UltimaLinha     As Integer
    Dim LinhaAtual      As Integer
    Dim Tarefa          As String
    Dim Periocidade     As String
    Dim PrazoFinal      As Double
    Dim DataConclusao   As Integer
    Dim Localidade      As String
    
    ' variaveis para planilha
    Dim AuxiliarNome    As Integer
    Dim ultimaPlanilha  As Integer
    Dim NomePlanilha    As String
    
    ' Variavel se inicia com 4:
    AuxiliarNome = 4
    
    ' Pega a linha da ultima planilha cadastrada:
    ultimaPlanilha = Sheets("TAREFAS").Range("F1048576").End(xlUp).Row
    
    ' Verifica se realmente existe uma planilha cadastrada:
    If (ultimaPlanilha <= 3) Then
    
        MsgBox "Por favor, Insira os nomes das planilhas na Planilha Tarefas."
        Exit Sub
    
    End If
    
    ' Pega a ultima linha não usada:
    UltimaLinha = Sheets("TAREFAS").Range("A1048576").End(xlUp).Row
    
    ' se nao tiver nada cadastrado nas tarefas:
    If (UltimaLinha = 4) Then
        Exit Sub
    Else ' se tiver algo cadastrado:
        LinhaAtual = 5
    End If
    
    ' laço para realizar em todas as planilhas:
    While (AuxiliarNome <= ultimaPlanilha)
        
        NomePlanilha = Sheets("TAREFAS").Range("F" & AuxiliarNome).Value
        
        ' Laço para percorrer todas as celulas:
        While (LinhaAtual <= UltimaLinha)
        
            ' Pega os valores das celulas:
            Tarefa = Sheets("TAREFAS").Range("A" & LinhaAtual).Value
            Periocidade = Sheets("TAREFAS").Range("B" & LinhaAtual).Value
            PrazoFinal = Sheets("TAREFAS").Range("C" & LinhaAtual).Value
            DataConclusao = Sheets("TAREFAS").Range("D" & LinhaAtual).Value
            
            ' retorna a posição da planilha:
            Localidade = Posicao(DataConclusao, NomePlanilha)
            
            ' Coloca no local correto:
            Sheets(NomePlanilha).Range(Localidade).Offset(1, 0).Value = Tarefa
            
            ' Adiciona mais 1 no contador:
            LinhaAtual = LinhaAtual + 1
                
        Wend
        
        ' Linha retorna a 0:
        LinhaAtual = 5
        
        ' Adiciona mais 1 na variavel auxiliar nome:
        AuxiliarNome = AuxiliarNome + 1
        
    Wend
    
End Sub

' função para encontrar a posição do determinado dia:
Function Posicao(DataFim As Integer, Planilha As String) As String

    Posicao = Achar(DataFim, Planilha)

End Function

Function Achar(Dia As Integer, Plan As String) As String

    ' Variaveis:
    Dim UltimaLinha     As Integer
    Dim Coluna          As String
    Dim auxiliarColuna  As Integer
    Dim LinhaAtual      As Integer
    Dim ValorLinha      As String
    Dim DiaString       As String
    
    ' Atribui o dia em uma string:
    DiaString = Dia
    
    ' Coloca 45 como a ultima Linha:
    UltimaLinha = 45
    
    ' A linha atual recebe 6:
    LinhaAtual = 6
    
    ' Coloca no auxiliar 0:
    auxiliarColuna = 0
    
    ' Um laço que percorre as colunas:
    While (auxiliarColuna <= 6)
        
        ' Chama a função para encontrar a coluna:
        Coluna = NovaColuna(auxiliarColuna)
        
        ' Laço para percorrer a planilha:
        While (LinhaAtual <= UltimaLinha)
        
            ' Pega o valor Contido na linha atual:
            ValorLinha = Sheets(Plan).Range(Coluna & LinhaAtual).Value
            
            ' Verifica se esta vazio o valor:
            If (ValorLinha <> "") Then
                ' Chama a função para converter o valor:
                ValorLinha = DiaConvertido(ValorLinha)
                
                ' Verifica se é o dia correto:
                If (ValorLinha = DiaString) Then
                    
                    ' Passar como retorno a coluna e a linha:
                    Achar = Coluna & LinhaAtual
                    
                    ' Sai da função:
                    Exit Function
                    
                End If
                 
            End If
            
            ' Adiciona mais 1 na variavel linha:
            LinhaAtual = LinhaAtual + 1
        
        Wend
        
        ' Linha Atual Retorna a 6
        LinhaAtual = 6
        
        ' Adiciona mais 1 na variavel:
        auxiliarColuna = auxiliarColuna + 1
        
    Wend

End Function

' função para converter o dia:
Function DiaConvertido(Data As String) As String

' Se por acaso retornar um erro:
On Error GoTo Oops
    
    ' Pega o dia da data:
    Dia = Day(Data)
    
    ' Retorna o dia:
    DiaConvertido = Dia
    
    ' Sai da função:
    Exit Function
    
Oops: ' Se der erro

    ' Retorna 0:
    DiaConvertido = 0
    
    Exit Function

End Function

' Função para mudar a coluna:
Function NovaColuna(Valor As Integer) As String

    ' Condições para mudar a coluna:
    If (Valor = 0) Then
    
        NovaColuna = "B"
    
    ElseIf (Valor = 1) Then
    
        NovaColuna = "D"
    
    ElseIf (Valor = 2) Then
    
        NovaColuna = "F"
    
    ElseIf (Valor = 3) Then
    
        NovaColuna = "H"
        
    ElseIf (Valor = 4) Then
    
        NovaColuna = "J"
        
    ElseIf (Valor = 5) Then
    
        NovaColuna = "L"
    
    ElseIf (Valor = 6) Then
    
        NovaColuna = "N"
        
    End If
    
End Function
