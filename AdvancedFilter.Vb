'---------------------------------------------
'                                            '
'   Nome: Miguel Fernandes Santos.           '
'   Data: 29/10/2021                         '
'   Projeto: Filtro Avançado.                '
'                                            '
'---------------------------------------------

Sub Filtrar_receber()
    
    ULTIMA_LINHA = Sheets("DADOS").Range("K1048576").End(xlUp).Row
    
    If ULTIMA_LINHA = 5 Then
    
        Exit Sub
        
    End If
    
    Set Base = Sheets("DADOS").Range("K5:P" & ULTIMA_LINHA)
    
    Set CRT = Sheets("DADOS").Range("K1:P2")
    
    Set COPIAR = Sheets("DADOS").Range("AK5:AP5")
    
    Base.AdvancedFilter xlFilterCopy, CRT, COPIAR
    
    ULTIMA_LINHA = Sheets("DADOS").Range("AK1048576").End(xlUp).Row
    
    If ULTIMA_LINHA = 5 Then
    
        ULTIMA_LINHA = 6
        
    End If
    
    CLIENTES_A_RECEBER.Listbox_Clientes.RowSource = "DADOS!AK6:AP" & ULTIMA_LINHA

End Sub
