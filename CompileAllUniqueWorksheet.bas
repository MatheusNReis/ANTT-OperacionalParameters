Sub CopiarDadosAtendimentosAPHeSocorroMecânico_dePastaLocal(ByVal Caminho As String)
    

'Compiles all concessioanire data into a single spreadsheet,
'with source data in sharepoint network, tested on local computer.
'Copy method can be chosen: line by line or column by column.

'Compila todos os dados de concessionárias numa única planilha,
'com dados de origem em sharepoint, testado em computador local.
'Método de cópia pode ser escolhido: linha por linha ou coluna por coluna.

'Created by Matheus Nunes Reis on 25/06/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: raw
'MIT License. Copyright © 2024 MatheusNReis


    Dim objWorkbook As Object, objWorksheet As Object
    Dim DestWorkbook As Workbook, DestWorksheet As Worksheet
    Dim i As Integer
    Dim LastRow As Long, LastRowObjWorksheet As Long
    
    
    ' Define a primeira planilha da pasta de trabalho de origem
    Set objWorkbook = Workbooks.Open(Caminho)
    Set objWorksheet = objWorkbook.Sheets(1)
    
    ' Define a pasta de trabalho de destino e a planilha de destino
    'Set DestWorkbook = Workbooks.Open("C:\Users\User.Name\Desktop\Atendimentos_APH_e_Socorro Mecânico.xlsm")
    Set DestWorksheet = ThisWorkbook.Sheets(1)
    
    ' Última linha com dados na planilha de origem
    LastRowObjWorksheet = objWorksheet.Cells(objWorksheet.Rows.Count, "A").End(xlUp).Row
    ' Última linha com dados na planilha de destino
    LastRowDest = DestWorksheet.Cells(DestWorksheet.Rows.Count, "A").End(xlUp).Row
    
    
    'Copiar dados coluna por coluna da planilha origem para a destino
    Dim lastcolumn As Long
    lastcolumn = objWorksheet.Cells(1, objWorksheet.Columns.Count).End(xlToLeft).Column
    
    For i = 1 To lastcolumn
        Dim rngOrigem As Range
        Dim rngDestino As Range

        ' Definir os intervalos de origem e destino
        Set rngOrigem = objWorksheet.Range(objWorksheet.Cells(2, i), objWorksheet.Cells(LastRowObjWorksheet, i))
        Set rngDestino = DestWorksheet.Cells(LastRowDest + 1, i)

        ' Copiar os valores
        rngDestino.Resize(rngOrigem.Rows.Count, rngOrigem.Columns.Count).Value = rngOrigem.Value
    Next i
    
    
    ' Copiar linha por linha - Processo funciona mais lentamente
    'For i = 1 To LastRowSharepoint
    '    objWorksheet.Range("A" & i & ":L" & i).Copy Destination:=DestWorksheet.Range("A" & i)
    'Next i
    ' Limpar a área de transferência
    'Application.CutCopyMode = False
    
    
    'Copia célula por célula - 'Processo funciona mais lentamente
    ' Copia os dados da primeira planilha para a planilha de destino
    'For j = 1 To 12 'Coluna L = 12
       'For i = 1 To LastRowSharepoint ' Copiar a quantidade de linhas da planilha sharepoint
            ' Copia os dados da coluna A da primeira planilha para a coluna A da planilha de destino
            'DestWorksheet.Cells(LastRowDest + i, j).Value = objWorksheet.Cells(i, j).Value
            ' Você pode expandir este loop para copiar mais colunas, se necessário
        'Next i
    'Next j
    
    ' Fecha a pasta de trabalho de origem
    objWorkbook.Close SaveChanges:=False
    
    
    MsgBox "Cópia completa para uma concessionária"
    
    
    ' Salva e fecha a pasta de trabalho de destino
    'ThisWorkbook.Save ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Salvar
    'DestWorkbook.Save
    'DestWorkbook.Close
    
    ' Libera a memória
    Set DestWorksheet = Nothing
    Set DestWorkbook = Nothing
    Set objWorksheet = Nothing
    Set objWorkbook = Nothing
    'objExcel.Quit
    'Set objExcel = Nothing
End Sub
