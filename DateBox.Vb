'---------------------------------------------
'                                            '
'   Nome: Miguel Fernandes Santos.           '
'   Data: 29/10/2021                         '
'   Projeto: Caixa de Texto Formato Data.    '
'                                            '
'---------------------------------------------

Dim comprovante1 As Boolean
Dim comprovante2 As Boolean
Dim comprovante3 As Boolean
Dim comprovante4 As Boolean

Private Sub txtdata_Change()

    If Len(txtdata.Text) = 2 And comprovante3 = True Then

        comprovante1 = False
        comprovante3 = False
        GoTo fim
        
    End If

    If Len(txtdata.Text) = 2 And comprovante1 = False Then

        comprovante1 = False
        comprovante3 = True
        txtdata.Text = txtdata.Text + "/"
        
    End If

    If Len(txtdata.Text) = 5 And comprovante4 = True Then

        comprovante2 = False
        comprovante4 = False
        GoTo fim
        
    End If

    If Len(txtdata.Text) = 5 And comprovante2 = False Then

        comprovante2 = False
        comprovante4 = True
        txtdata.Text = txtdata.Text + "/"
        
    End If

    fim:

End Sub
