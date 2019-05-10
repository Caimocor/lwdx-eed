'Sub teste()
'    Dim ie As InternetExplorer
'    Dim webpage As HTMLDocument
'    Set ie = New InternetExplorer
'    ie.Visible = True
'    ie.Navigate ("https://en.wikipedia.org/wiki/List_of_countries_and_dependencies_by_population")
'    Do While ie.Busy = True
'    Loop
'
'    Set webpage = ie.Document
'    table_data = webpage.getElementsByTagName("tr")
'
'    Debug.Print table_data.innerText
'End Sub

'Sub teste()
'    Dim ie As InternetExplorer
'    Dim webpage As HTMLDocument
'    Set ie = New InternetExplorer
'    ie.Visible = True
'    ie.Navigate ("https://www.submarino.com.br/?origem_clique=banner")
'    Do While ie.Busy = True
'    Loop
'
'    Set webpage = ie.Document
'    table_data = webpage.getElementsByClassName("lst-lnk rp")
'    Debug.Print "Caiooo"
'    Debug.Print table_data.innerText
'End Sub

Sub teste()
    'Inicia o navegador
    Dim ie As InternetExplorer
    Dim webpage As HTMLDocument
    Set ie = New InternetExplorer
    ie.Visible = True
    
    
    
    RowCount = 0
    Set sh = ActiveSheet
    
    'Inicia o processo de verificar linha a linha
    For Each rw In sh.Rows
      'Verifica se não existem mais linhas
      If sh.Cells(rw.Row, 1).Value = "" Then
        Exit For
      End If
      RowCount = RowCount + 1
      
      'caso exista, pega o valor da linha e coloca no link do navegador
      ie.Navigate ("https://www.submarino.com.br/?origem_clique=banner")
      Do While ie.Busy = True
      Loop
      
      Set webpage = ie.Document
      table_data = webpage.getElementsByClassName("lst-lnk rp")
      Debug.Print "----------------------------------------------"
      
      
    Next rw

End Sub
