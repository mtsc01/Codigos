Sub nomeCurso()
    
    Dim ultimalinha As Integer
    Dim x As Range
    
    ultimalinha = Cells(Rows.Count, 1).End(xlUp).Row
    
    For Each celula In Plan9.Range("a2:a" & ultimalinha)
        For Each plan In Worksheets
            Set x = plan.Cells.Find(celula.Value)
            If Not x Is Nothing Then
                'celula.Offset(0, 1).Value = plan.Name
                celula.Offset(0, 2).Value = x.Offset(0, 1).Value
                celula.Offset(0, 3).Value = x.Offset(0, 2).Value
                celula.Offset(0, 4).Value = x.Offset(0, 23).Value
                celula.Offset(0, 5).Value = x.Offset(0, 24).Value
                celula.Offset(0, 6).Value = x.Offset(0, 26).Value
                celula.Offset(0, 7).Value = x.Offset(0, 27).Value
                celula.Offset(0, 8).Value = x.Offset(0, 28).Value
                celula.Offset(0, 9).Value = x.Offset(0, 29).Value
                celula.Offset(0, 10).Value = x.Offset(0, 30).Value
                celula.Offset(0, 11).Value = x.Offset(0, 18).Value
                celula.Offset(0, 12).Value = x.Offset(0, 43).Value
                celula.Offset(0, 13).Value = x.Offset(0, 44).Value
                Exit For
            End If
        Next
    Next
    
End Sub
