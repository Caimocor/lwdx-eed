Sub Macro()

    Set sh = ActiveWorkbook.Sheets("Plan1")

    For r = 2 To 100
    
        'Testa para o último caso da mesa
        If IsEmpty(sh.Cells(1, r).Value) Then
            sh.Cells(1, r).Value = "Desk " & sh.Cells(1, r - 1).Value
            Exit For
        End If
    
        'Procura áreas com o mesmo nome e insere uma coluna entre elas que será a Desk
        'Compara se os campos estão em branco
        If Not (IsEmpty(sh.Cells(1, r).Value)) And Not (IsEmpty(sh.Cells(1, r + 1).Value)) Then
        
            'compara se as mesas possuem nomes parecidos
            If StrComp(sh.Cells(1, r).Value, sh.Cells(1, r + 1).Value, vbTextCompare) = -1 Then
                
                Columns(r + 1).Insert Shift = xlToRight
                sh.Cells(1, r + 1).Value = "Desk " & sh.Cells(1, r).Value
            End If
            

        
        End If
    Next r

    Set sh_heads = ActiveWorkbook.Sheets("Heads")
    'Determina quem são os Heads de cada área
    head_area_1 = sh_heads.Cells(1, 2).Value
    head_area_2 = sh_heads.Cells(2, 2).Value
    head_area_3 = sh_heads.Cells(3, 2).Value
    For r = 2 To 100

        'Verifica se o Head é o que está sendo percorrido, caso seja, pinta de azul
        If StrComp(sh.Cells(2, r).Value, head_area_1, vbTextCompare) = 0 And (StrComp(sh.Cells(1, r).Value, "área_1", vbTextCompare) = 0) Then
            sh.Cells(2, r).Interior.ColorIndex = 37
            Debug.Print "testee"
        End If

        'Verifica se o Head é o que está sendo percorrido, caso seja, pinta de azul
        If StrComp(sh.Cells(2, r).Value, head_area_2, vbTextCompare) = 0 And (StrComp(sh.Cells(1, r).Value, "área_2", vbTextCompare) = 0) Then
            sh.Cells(2, r).Interior.ColorIndex = 37
            Debug.Print "testee"
        End If

        'Verifica se o Head é o que está sendo percorrido, caso seja, pinta de azul
        If StrComp(sh.Cells(2, r).Value, head_area_3, vbTextCompare) = 0 And (StrComp(sh.Cells(1, r).Value, "área_3", vbTextCompare) = 0) Then
            sh.Cells(2, r).Interior.ColorIndex = 37
            Debug.Print "testee"
        End If
    Next r
    

End Sub


