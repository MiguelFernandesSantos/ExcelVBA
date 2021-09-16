Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

If Application.Visible = False Then
    
    If CloseMode = 0 Then
        
        CloseMode = 1
        
             Dim Pergunta As VbMsgBoxResult
             
             Pergunta = MsgBox("Deseja salvar antes de fechar o formulario?", vbYesNo)
             
             If Pergunta = vbYes Then
             
                 ThisWorkbook.Save
                 ThisWorkbook.Close
                
             Else
                 
                 Application.DisplayAlerts = False
                 ThisWorkbook.Close
                
             End If
             
    End If
    
End If

End Sub