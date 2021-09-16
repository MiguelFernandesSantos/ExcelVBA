Sub Atualizar_Tabelas_a_pagar()

    Sheets("Tabelas").PivotTables("Tabela dinâmica1").PivotCache.Refresh
    Sheets("Tabelas").PivotTables("Tabela dinâmica2").PivotCache.Refresh
    Sheets("Tabelas").PivotTables("Tabela dinâmica3").PivotCache.Refresh
    
End Sub