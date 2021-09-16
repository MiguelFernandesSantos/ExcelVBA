Dim virgula As Boolean

Private Sub txtvalor_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 44 Or virgula = True) Then
        
        KeyAscii = 0
        
    End If
    
    If (KeyAscii = 44) Then
    
        virgula = True
        
    End If
    
    If (txtvalor.TextLength = 0) Then
    
        virgula = False
        
    End If
    
    
    If (txtvalor.TextLength = 0 And KeyAscii = 44) Then
        
        txtvalor.Text = "0,"
        virgula = True
        KeyAscii = 0
        
    End If
    
End Sub