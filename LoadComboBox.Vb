'---------------------------------------------
'                                            '
'   Nome: Miguel Fernandes Santos.           '
'   Data: 29/10/2021                         '
'   Projeto: Carregar Combo Box com Laço.    '
'                                            '
'---------------------------------------------

Private Sub UserForm_Initialize()

linha_fim = Sheets("DADOS").Range("AF1048576").End(xlUp).Row

txtdescricao.Clear

For linha1 = 1 To linha_fim

    txtdescricao.AddItem DADOS.Cells(linha1, 1)
    
Next linha1

End Sub
