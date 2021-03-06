Sub NovoCalendario()
    
    ' Variaveis para erros:
    Dim erroLinha1      As Integer
    Dim erroColuna1     As Integer
    Dim TipoErro        As Integer
    
    ' variaveis para tarefas
    Dim UltimaLinha     As Integer
    Dim LinhaAtual      As Integer
    Dim Tarefa          As String
    Dim Periocidade     As String
    Dim DataConclusao   As Integer
    Dim Localidade      As String
    Dim EstaVazio       As String
    Dim quantidade      As Integer
    Dim Adicionado      As Boolean
    
    ' variaveis para planilha
    Dim AuxiliarNome    As Integer
    Dim ultimaPlanilha  As Integer
    Dim NomePlanilha    As String
    Dim QtsDias         As Integer
    Dim ProximoDia      As String
    Dim qtsLinhas       As Integer
    Dim LinhaUm         As Integer
    Dim LinhaDois       As Integer
    Dim NovaUltimaLinha As Integer
     
    ' Variaveis para adicionar a linha:
    Dim LocalUm         As String
    Dim localDois       As String
    Dim ColunaUm        As String
    Dim ColunaDois      As String

    ' Quantidade se inicia em 1:
    quantidade = 1
    
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
    
        MsgBox "Por favor, insira as tarefas que você deseja adicionar."
        Exit Sub
        
    Else ' se tiver algo cadastrado:
    
        LinhaAtual = 5
        
    End If
    
    ' laço para realizar em todas as planilhas:
    While (AuxiliarNome <= ultimaPlanilha)
        
        ' Verificia se existe uma tarefa para ser cadastrada:
        If (Sheets("TAREFAS").Range("F" & AuxiliarNome).Value <> "") Then
        
            NomePlanilha = Sheets("TAREFAS").Range("F" & AuxiliarNome).Value
            
        Else ' se não
        
            erroLinha1 = Sheets("TAREFAS").Range("F" & AuxiliarNome).Row
            MsgBox "Erro1: Não existe uma planilha cadastrada na linha " & erroLinha1 & "."
            Exit Sub
            
        End If
        
        ' Verifica se existe dias cadastrado:
        If (Sheets("TAREFAS").Range("F" & AuxiliarNome).Offset(0, 1).Value <> "") Then
            
            'On Error GoTo Erro3
            QtsDias = Sheets("TAREFAS").Range("F" & AuxiliarNome).Offset(0, 1).Value
            
            
        Else ' Se não:
        
            erroLinha1 = Sheets("TAREFAS").Range("F" & AuxiliarNome).Offset(0, 1).Row
            MsgBox "Erro2: Não existe a quantidade de dias da planilha na linha " & erroLinha1 & "."
            Exit Sub
        
        End If
        
        ' Adiciona mais 1 na variavel auxiliar nome:
        AuxiliarNome = AuxiliarNome + 1
        
        ' Laço para percorrer todas as celulas:
        While (LinhaAtual <= UltimaLinha)
        
            ' verifica se todos os campos possuem valores:
            If (Sheets("TAREFAS").Range("A" & LinhaAtual).Value <> "" And Sheets("TAREFAS").Range("B" & LinhaAtual).Value <> "" And Sheets("TAREFAS").Range("D" & LinhaAtual).Value <> "") Then
                
                ' Pega os valores das celulas:
                Tarefa = Sheets("TAREFAS").Range("A" & LinhaAtual).Value
                Periocidade = Sheets("TAREFAS").Range("B" & LinhaAtual).Value
                
                On Error GoTo erro6
                DataConclusao = Sheets("TAREFAS").Range("D" & LinhaAtual).Value
                
            Else ' se não:
            
                erroLinha1 = Sheets("TAREFAS").Range("A" & LinhaAtual).Row
                MsgBox "Erro5: Não foi possivel encontrar todas as informações necessarias na tarefa da linha " & erroLinha1 & "."
                Exit Sub
                
            End If
            
            ' se o mes nao tiver a quantidade de dias necessarios:
            If (DataConclusao > QtsDias) Then
                
                ' Coloca no ultimo dia do mes:
                DataConclusao = QtsDias
            
            End If
            
            ' retorna a posição da planilha:
            Localidade = Posicao(DataConclusao, NomePlanilha)
              
            ' Se a função retornar um erro:
            If (Localidade = "Erro4") Then
                
                MsgBox "Erro4: Dificuldade na localização do dia " & DataConclusao & " verifique se o dia esta correto, ou se existe X abaixo do calendario.  Planilha: " & NomePlanilha
                Exit Sub
            
            End If
              
            If ((DataConclusao + 7) >= QtsDias) Then
                
                ' Encontra o local do ultimo dia:
                ProximoDia = Posicao(QtsDias, NomePlanilha)
                
                If (Localidade = "Erro4" Or ProximoDia = "Erro4") Then
                
                    MsgBox "Erro4: Dificuldade na localização do dia " & QtsDias & " verifique se o dia esta correto, ou se existe X abaixo do calendario.   Planilha: " & NomePlanilha
                    Exit Sub
            
                End If
                
                ' Pega a coluna do dia da tarefa:
                LinhaUm = Sheets(NomePlanilha).Range(Localidade).Row
                
                ' Pega a coluna do ultimo dia:
                LinhaDois = Sheets(NomePlanilha).Range(ProximoDia).Row
                
                ' Se O dia procurado ser na mesma linha:
                If (LinhaUm = LinhaDois) Then
                
                    NovaUltimaLinha = Sheets(NomePlanilha).Range("B1048576").End(xlUp).Row
                    NovaUltimaLinha = Sheets(NomePlanilha).Range("B" & NovaUltimaLinha).End(xlUp).Row - 1
                    
                    qtsLinhas = NovaUltimaLinha - LinhaUm
                    
                Else
                    
                    ' Descobre quantas linhas livres tem:
                    qtsLinhas = DiferencaNormal(Localidade, ProximoDia, NomePlanilha)
                    
                    If (qtsLinhas = -2) Then
                
                        MsgBox "Erro6: Verifique se abaixo do calendário  existe duas linhas com X."
                        Exit Sub
                
                    End If
                
                End If
            
            Else
                ' Descobre o proximo dia:
                ProximoDia = Posicao((DataConclusao + 7), NomePlanilha)
                
                '
                qtsLinhas = DiferencaNormal(Localidade, ProximoDia, NomePlanilha)
                
                If (qtsLinhas = -2) Then
                
                    MsgBox "Erro6: Verifique se abaixo do calendário existe duas linhas com X"
                    Exit Sub
                
                End If
                
            End If
              
            ' Laço para encontrar um valor livre:
            While (Adicionado = False)
                
                ' pega o valor da planilha
                EstaVazio = Sheets(NomePlanilha).Range(Localidade).Offset(quantidade, 0).Value
                
                ' Verifica se tem algo:
                If (EstaVazio = "" Or EstaVazio = " ") Then
                
                    ' Coloca no local correto:
                    Sheets(NomePlanilha).Range(Localidade).Offset(quantidade, 0).Value = Tarefa
                    Adicionado = True
                    
                Else ' Se estiver:
                    
                    ' adiciona mais 1 na variavel para se deslocar para baixo:
                    quantidade = quantidade + 1
                    
                    ' Verifica se ja foi cadastrada:
                    If (EstaVazio = Tarefa) Then
                        
                        ' Sai do laço
                        Adicionado = True
                        GoTo FImLaco
                        
                    End If
                
                End If
                              
                ' Verifica se existe espaço para cadastrar:
                If (quantidade > qtsLinhas) Then
                    
                    ' Adiciona uma nova Linha:
                    Sheets(NomePlanilha).Range(Localidade).Offset(quantidade - 1, 0).EntireRow.Insert
                    
                    ' Pega o range da nova celula:
                    ColunaUm = Sheets(NomePlanilha).Range(Localidade).Offset(quantidade - 1, 0).Column
                    ColunaUm = QualColuna(CInt(ColunaUm))
                    LinhaUm = Sheets(NomePlanilha).Range(Localidade).Offset(quantidade - 1, 0).Row
                    LocalUm = ColunaUm & LinhaUm
                    
                    ' Pega o range da nova celula
                    ColunaDois = Sheets(NomePlanilha).Range(Localidade).Offset(quantidade - 1, 1).Column
                    ColunaDois = QualColuna(CInt(ColunaDois))
                    LinhaDois = Sheets(NomePlanilha).Range(Localidade).Offset(quantidade - 1, 0).Row
                    localDois = ColunaDois & LinhaDois
                    
                    ' Mescla as celulas:
                    'Sheets(NomePlanilha).Range(LocalUm & ":" & localDois).Merge
                    Sheets(NomePlanilha).Range("B" & LinhaUm & ":" & "C" & LinhaUm).Merge
                    Sheets(NomePlanilha).Range("D" & LinhaUm & ":" & "E" & LinhaUm).Merge
                    Sheets(NomePlanilha).Range("F" & LinhaUm & ":" & "G" & LinhaUm).Merge
                    Sheets(NomePlanilha).Range("H" & LinhaUm & ":" & "I" & LinhaUm).Merge
                    Sheets(NomePlanilha).Range("J" & LinhaUm & ":" & "K" & LinhaUm).Merge
                    Sheets(NomePlanilha).Range("L" & LinhaUm & ":" & "M" & LinhaUm).Merge
                    Sheets(NomePlanilha).Range("N" & LinhaUm & ":" & "O" & LinhaUm).Merge
                    
                    ' Adiciona a tarefa:
                    
                    Sheets(NomePlanilha).Range(Localidade).Offset(quantidade, 0).Value = Tarefa
                    
                    ' Adiciona mais 1 na quantidade de linhas:
                    qtsLinhas = qtsLinhas + 1
                    
                    ' Fala que ja foi adicionado
                    Adicionado = True
                    
                End If
                
FImLaco:
                
            Wend
            
            ' Quantidade retorna a 1:
            quantidade = 1
            
            ' Adiciona mais 1 no contador:
            LinhaAtual = LinhaAtual + 1
            
            Adicionado = False
                
        Wend
        
        ' Linha retorna a 0:
        LinhaAtual = 5
        
    Wend
    
Exit Sub

' Oops:
 '   MsgBox ("Verifique se existe uma planilha com o nome ou se existe tarefas cadastradas")
    
Erro3:
    
    erroLinha1 = Sheets("TAREFAS").Range("F" & AuxiliarNome).Offset(0, 1).Row
    MsgBox "Erro3: A quantidade de dias inseridos da planilha na linha " & erroLinha1 & " não é um numero valido."

erro6:
    
erroLinha1 = Sheets("TAREFAS").Range("D" & LinhaAtual).Row
MsgBox "Erro3: O dia para realizar a tarefa da linha " & erroLinha1 & " não é um dia valido."
    

End Sub

' função para encontrar a posição do determinado dia:
Function Posicao(DataFim As Integer, planilha As String) As String

    Posicao = Achar(DataFim, planilha)

End Function
Function Achar(Dia As Integer, Plan As String) As String

On Error GoTo Oops

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
    UltimaLinha = Sheets(Plan).Range("B1048576").End(xlUp).Row
    
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
    
If (Achar = "") Then

    Achar = "Erro4"

End If

Exit Function

Oops:

    Achar = "Erro4"
    
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
' Função para saber quantas linhas livres tem para cadastrar:
Function DiferencaNormal(PrimeiroDia As String, SegundoDia As String, NomePlanilha As String) As Integer

' Se por acaso retornar um erro:
On Error GoTo Oops

    ' Variaveis:
    Dim LinhaPrimeiroDia As Integer
    Dim LinhaSegundoDia  As Integer
    
    ' Pega a linha da primeira:
    LinhaPrimeiroDia = Sheets(NomePlanilha).Range(PrimeiroDia).Row
    
    ' Pega a linha da Segunda:
    LinhaSegundoDia = Sheets(NomePlanilha).Range(SegundoDia).Row
    
    ' Faz a diferença e tambem retorna:
    DiferencaNormal = (LinhaSegundoDia - LinhaPrimeiroDia) - 1
    
Exit Function
    
Oops:

DiferencaNormal = -2

End Function
' Função para mudar a coluna:
Function QualColuna(Valor As Integer) As String

    ' Condições para mudar a coluna:
    If (Valor = 2) Then
    
        QualColuna = "B"
    
    ElseIf (Valor = 3) Then
    
        QualColuna = "C"
    
    ElseIf (Valor = 4) Then
    
        QualColuna = "D"
    
    ElseIf (Valor = 5) Then
    
        QualColuna = "E"
        
    ElseIf (Valor = 6) Then
    
        QualColuna = "F"
        
    ElseIf (Valor = 7) Then
    
        QualColuna = "G"
    
    ElseIf (Valor = 8) Then
    
        QualColuna = "H"
    
    ElseIf (Valor = 9) Then
    
        QualColuna = "I"
        
    ElseIf (Valor = 10) Then
    
        QualColuna = "J"
        
    ElseIf (Valor = 11) Then
    
        QualColuna = "K"
    
    ElseIf (Valor = 12) Then
    
        QualColuna = "L"
        
    ElseIf (Valor = 13) Then
    
        QualColuna = "M"
        
    ElseIf (Valor = 14) Then
    
        QualColuna = "N"
    
    ElseIf (Valor = 15) Then
    
        QualColuna = "O"
        
    End If
    
End Function
