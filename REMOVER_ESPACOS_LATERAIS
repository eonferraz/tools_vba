Sub REMOVER_ESPACOS_LATERAIS()

    Dim celula As Range
    
    For Each celula In Selection
    
            While Left(celula.Value, 1) = " "
                
                celula.Value = Right(celula.Value, Len(celula.Value) - 1)
            
            Wend
        
            While Right(celula.Value, 1) = " "
                
                celula.Value = Left(celula.Value, Len(celula.Value) - 1)
            
            Wend
            
    Next celula

End Sub



